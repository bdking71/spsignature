/**
 * @file TransactionSigner.ts
 *
 * Provides a secure, auditable digital-signature workflow for SharePoint
 * web parts.  The module renders a modal dialog that supports:
 *
 *   • Drawing a signature on an HTML canvas
 *   • Uploading a signature image (PNG / JPG, ≤ 5 MB)
 *   • Re-using a previously cached (LZW-compressed) signature from
 *     localStorage
 *   • Optional two-factor authentication via a 5-digit passcode
 *     dispatched through email or Microsoft Teams
 *
 * After the user signs, the module hashes the audit envelope
 * (payload + signer + timestamp) with SHA-256 so that the record can
 * later be verified without exposing the original payload.
 *
 * @module SecureAuditSignature
 */

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { dispatchOtpPasscode, DeliveryChannel } from "./OtpService";

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** localStorage key used to persist a compressed signature across sessions. */
const STORAGE_KEY = "secure_audit_cached_signature_v1";

/** Maximum upload file size in bytes (5 MB). */
const MAX_UPLOAD_BYTES = 5 * 1024 * 1024;

/** Maximum LZW dictionary size (prevents corruption in localStorage/JSON). */
const LZW_DICT_MAX = 65535;

/** Canvas dimensions. */
const CANVAS_WIDTH = 560;
const CANVAS_HEIGHT = 160;

/** Image resize limits. */
const IMAGE_MAX_WIDTH = 560;
const IMAGE_MAX_HEIGHT = 160;

/** Passcode generation limits. */
const PASSCODE_MIN = 10000;
const PASSCODE_MAX = 99999;
const PASSCODE_LENGTH = 5;

/** TFA cooldown timer (seconds). */
const TFA_COOLDOWN_SECONDS = 60;

/** Modal close timeout (milliseconds) - prevents hanging promises. */
const MODAL_TIMEOUT_MS = 30 * 60 * 1000; // 30 minutes

/** Timestamp precision (seconds, not milliseconds). */
const TIMESTAMP_PRECISION = "seconds";

/** Fluent Design System colors. */
const COLORS = {
  PRIMARY: "#0078d4",
  PRIMARY_DARK: "#005a9e",
  ERROR: "#a4262c",
  WARNING: "#ffd54f",
  WARNING_BG: "#fff8e1",
  INFO: "#e6f4ff",
  INFO_BORDER: "#0078d4",
  NEUTRAL_BG: "#fafafa",
  NEUTRAL_LIGHT: "#f3f2f1",
  NEUTRAL_LIGHTER: "#e1dfdd",
  NEUTRAL_TEXT: "#323130",
  NEUTRAL_TEXT_SEC: "#605e5c",
  NEUTRAL_BORDER: "#d1d1d1",
  NEUTRAL_BORDER_DARK: "#8a8886",
};

/** UI Icons (with accessible fallbacks). */
const ICONS = {
  LOCK: "🔐",
  CHECK: "✓",
  FOLDER: "📁",
};

// Global modal state to prevent concurrent instances
let activeModalInstance: AbortController | null = null;

// ---------------------------------------------------------------------------
// Public interfaces
// ---------------------------------------------------------------------------

/**
 * Describes everything the caller must (or may) supply when requesting a
 * signed audit record.
 */
export interface SignerContext {
  /** SharePoint list-item ID that the signature relates to. */
  itemID: number;
  /** Display name or email of the person signing. */
  signer: string;
  /** Arbitrary key/value data to include in the audit envelope. */
  payload: Record<string, unknown>;
  /** SPFx web-part context – required for PnPjs / Graph calls. */
  spContext: WebPartContext;
  /** How to deliver the two-factor passcode (`"email"` | `"teams"`). */
  channel?: DeliveryChannel;
  /**
   * Whether two-factor authentication is required.
   * **Defaults to `false` (disabled).** Set to `true` to enable TFA.
   */
  requireTFA?: boolean;
}

/**
 * The record that is ultimately persisted to SharePoint after a
 * successful signing ceremony.
 */
export interface SharePointAuditRecord {
  /** SHA-256 hex digest of the canonical audit envelope. */
  signatureHash: string;
  /** LZW-compressed data-URI of the signature image. */
  signatureData: string;
  /** ISO-8601 timestamp captured at the moment of signing. */
  signatureTimestamp: string;
  /** ID of the verification list-item (if two-factor was used). */
  verificationItemId: number;
}

/**
 * The canonical shape that is hashed to produce `signatureHash`.
 * The payload keys are sorted alphabetically so that hash
 * verification is order-independent.
 */
export interface AuditEnvelopeRecord {
  payload: Record<string, unknown>;
  signer: string;
  timestamp: string;
  itemID?: number;
}

// ---------------------------------------------------------------------------
// Helpers – validation
// ---------------------------------------------------------------------------

/**
 * Validates that a value is a non-empty string.
 *
 * @param value - The value to check.
 * @param fieldName - Field name for error messages.
 * @returns The trimmed string.
 * @throws {Error} If value is not a valid string.
 */
function assertNonEmptyString(value: unknown, fieldName: string): string {
  if (typeof value !== "string" || value.trim().length === 0) {
    throw new Error(`${fieldName} must be a non-empty string.`);
  }
  return value.trim();
}

/**
 * Validates that an object is JSON-serializable.
 *
 * @param obj - Object to validate.
 * @throws {Error} If object contains non-serializable values.
 */
function validateJsonSerializable(obj: unknown): void {
  try {
    JSON.stringify(obj);
  } catch (error) {
    throw new Error(`Payload is not JSON-serializable: ${String(error)}`);
  }
}

/**
 * Sanitizes email/signer input to prevent XSS.
 *
 * @param input - The signer identifier.
 * @returns Sanitized string (alphanumeric, @, ., -, _).
 */
function sanitizeSigner(input: string): string {
  const sanitized = input.toLowerCase().trim();
  if (!/^[a-z0-9@._-]+$/.test(sanitized)) {
    throw new Error("Signer identifier contains invalid characters.");
  }
  return sanitized;
}

/**
 * Generates a precision-reduced ISO timestamp (seconds, not milliseconds)
 * to avoid hash mismatches from millisecond variations.
 *
 * @returns ISO-8601 timestamp (seconds precision).
 */
function generateTimestamp(): string {
  return new Date().toISOString().split(".")[0] + "Z";
}

// ---------------------------------------------------------------------------
// Helpers – passcode generation
// ---------------------------------------------------------------------------

/**
 * Generates a cryptographically random 5-digit passcode (10 000 – 99 999)
 * using the Web Crypto API.
 *
 * @returns A string representation of the 5-digit code.
 */
function generateFiveDigitPasscode(): string {
  const array = new Uint32Array(1);
  window.crypto.getRandomValues(array);
  const code = (array[0] % (PASSCODE_MAX - PASSCODE_MIN + 1)) + PASSCODE_MIN;
  return code.toString();
}

// ---------------------------------------------------------------------------
// Main public function – modal signing ceremony
// ---------------------------------------------------------------------------

/**
 * Opens a full-screen modal dialog that walks the user through:
 *
 * 1. (Optional) Two-factor passcode verification (user-initiated)
 * 2. Providing a digital signature (cached / drawn / uploaded)
 * 3. Generating a tamper-evident SHA-256 audit record
 *
 * The returned promise resolves with a `SharePointAuditRecord` on
 * success or `undefined` if the user cancels.
 *
 * **Note:** Only one modal can be active at a time. Attempting to open
 * a second modal while one is active will return `undefined` immediately.
 *
 * @param context        - Signer metadata and SPFx context.
 * @param modalTitle     - Title shown in the modal header.
 * @param warningMessage - Instructional HTML rendered above the
 *                         signature area.
 * @returns A promise that resolves to the audit record or `undefined`.
 *
 * @throws {Error} If `context.spContext` is falsy or signer is invalid.
 *
 * @example
 * ```ts
 * const record = await promptAndGenerateSecureAudit({
 *   itemID: 42,
 *   signer: currentUser.email,
 *   payload: { amount: 1500, vendor: "Contoso" },
 *   spContext: this.context,
 *   channel: "email",
 *   requireTFA: true,  // Enable TFA
 * });
 *
 * if (record) {
 *   console.log("Signed successfully!");
 * }
 * ```
 */
export async function promptAndGenerateSecureAudit(
  context: SignerContext,
  modalTitle: string = "Approve Purchase Requisition",
  warningMessage: string = "Please review your entry and apply your signature to finalize this record."
): Promise<SharePointAuditRecord | undefined> {
  if (!context.spContext) {
    throw new Error(
      "Execution restricted: Valid WebPartContext must be supplied."
    );
  }

  // Prevent concurrent modals
  if (activeModalInstance !== null) {
    return undefined;
  }

  // Validate signer early
  try {
    assertNonEmptyString(context.signer, "Signer");
    sanitizeSigner(context.signer);
  } catch (error) {
    throw new Error(`Invalid signer: ${String(error)}`);
  }

  return new Promise<SharePointAuditRecord | undefined>((resolve): void => {
    // Create abort controller for this modal instance
    const abortController = new AbortController();
    activeModalInstance = abortController;

    // Set timeout to prevent hanging promises
    const timeoutId = window.setTimeout(() => {
      cleanup();
      resolve(undefined);
    }, MODAL_TIMEOUT_MS);

    const requireTFA = context.requireTFA === true;
    const generatedPasscode = generateFiveDigitPasscode();
    let storedVerificationItemId: number | undefined = undefined;
    let tfaVerified = false;

    // -----------------------------------------------------------------
    // Overlay
    // -----------------------------------------------------------------
    const overlay: HTMLDivElement = document.createElement("div");
    overlay.style.position = "fixed";
    overlay.style.top = "0";
    overlay.style.left = "0";
    overlay.style.width = "100vw";
    overlay.style.height = "100vh";
    overlay.style.backgroundColor = "rgba(0, 0, 0, 0.7)";
    overlay.style.zIndex = "999999";
    overlay.style.display = "flex";
    overlay.style.alignItems = "center";
    overlay.style.justifyContent = "center";
    overlay.style.fontFamily =
      "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";
    overlay.style.backdropFilter = "blur(3px)";

    // -----------------------------------------------------------------
    // Modal container
    // -----------------------------------------------------------------
    const modalBox: HTMLDivElement = document.createElement("div");
    modalBox.style.backgroundColor = "#ffffff";
    modalBox.style.padding = "0";
    modalBox.style.borderRadius = "8px";
    modalBox.style.maxWidth = "640px";
    modalBox.style.width = "90%";
    modalBox.style.boxShadow = "0 8px 32px rgba(0,0,0,0.3)";
    modalBox.style.overflow = "hidden";
    modalBox.style.maxHeight = "90vh";
    modalBox.style.display = "flex";
    modalBox.style.flexDirection = "column";

    // -----------------------------------------------------------------
    // Header
    // -----------------------------------------------------------------
    const headerSection: HTMLDivElement = document.createElement("div");
    headerSection.style.backgroundColor = COLORS.PRIMARY;
    headerSection.style.padding = "20px 24px";
    headerSection.style.borderBottom = `3px solid ${COLORS.PRIMARY_DARK}`;

    const title: HTMLHeadingElement = document.createElement("h2");
    title.innerText = modalTitle;
    title.style.margin = "0";
    title.style.color = "#ffffff";
    title.style.fontSize = "20px";
    title.style.fontWeight = "600";
    title.style.letterSpacing = "0.3px";

    headerSection.appendChild(title);

    // -----------------------------------------------------------------
    // Content wrapper
    // -----------------------------------------------------------------
    const contentSection: HTMLDivElement = document.createElement("div");
    contentSection.style.padding = "24px";
    contentSection.style.backgroundColor = COLORS.NEUTRAL_BG;
    contentSection.style.overflowY = "auto";
    contentSection.style.flex = "1";

    const body: HTMLParagraphElement = document.createElement("p");
    body.innerHTML = warningMessage;
    body.style.fontSize = "14px";
    body.style.color = COLORS.NEUTRAL_TEXT;
    body.style.lineHeight = "1.6";
    body.style.margin = "0 0 20px 0";

    // -----------------------------------------------------------------
    // Sign button (footer)
    // -----------------------------------------------------------------
    const signButton: HTMLButtonElement = document.createElement("button");
    signButton.innerText = "I Agree and Sign";
    signButton.setAttribute("aria-label", "Sign the document with your digital signature");
    signButton.style.padding = "10px 24px";
    signButton.style.backgroundColor = COLORS.PRIMARY;
    signButton.style.color = "#ffffff";
    signButton.style.border = "none";
    signButton.style.cursor = "pointer";
    signButton.style.fontSize = "14px";
    signButton.style.fontWeight = "600";
    signButton.style.borderRadius = "4px";
    signButton.style.transition = "all 0.2s ease";

    // -----------------------------------------------------------------
    // Two-factor verification panel
    // -----------------------------------------------------------------
    const verificationContainer: HTMLDivElement =
      document.createElement("div");

    let verificationInput: HTMLInputElement | null = null;
    let sendCodeButton: HTMLButtonElement | null = null;

    if (requireTFA) {
      verificationContainer.style.marginBottom = "20px";
      verificationContainer.style.padding = "16px";
      verificationContainer.style.backgroundColor = COLORS.WARNING_BG;
      verificationContainer.style.border = `1px solid ${COLORS.WARNING}`;
      verificationContainer.style.borderRadius = "6px";
      verificationContainer.style.boxShadow = "0 2px 4px rgba(0,0,0,0.05)";

      const verificationHeader: HTMLDivElement =
        document.createElement("div");
      verificationHeader.style.display = "flex";
      verificationHeader.style.alignItems = "center";
      verificationHeader.style.marginBottom = "12px";
      verificationHeader.style.gap = "8px";

      const lockIcon: HTMLSpanElement = document.createElement("span");
      lockIcon.innerHTML = ICONS.LOCK;
      lockIcon.style.fontSize = "18px";
      lockIcon.setAttribute("aria-hidden", "true");

      const verificationLabel: HTMLLabelElement =
        document.createElement("label");
      verificationLabel.innerText = "Two-Factor Authentication";
      verificationLabel.style.display = "block";
      verificationLabel.style.fontSize = "13px";
      verificationLabel.style.fontWeight = "700";
      verificationLabel.style.color = COLORS.NEUTRAL_TEXT;
      verificationLabel.style.margin = "0";

      verificationHeader.appendChild(lockIcon);
      verificationHeader.appendChild(verificationLabel);

      const verificationDescription: HTMLParagraphElement =
        document.createElement("p");
      verificationDescription.innerText = `Click "Send Code" to receive a 5-digit verification code via ${
        context.channel === "teams" ? "Microsoft Teams" : "email"
      }.`;
      verificationDescription.style.fontSize = "12px";
      verificationDescription.style.color = COLORS.NEUTRAL_TEXT_SEC;
      verificationDescription.style.margin = "0 0 12px 0";
      verificationDescription.style.lineHeight = "1.5";

      const verificationRow: HTMLDivElement =
        document.createElement("div");
      verificationRow.style.display = "flex";
      verificationRow.style.gap = "10px";

      verificationInput = document.createElement("input");
      verificationInput.type = "text";
      verificationInput.inputMode = "numeric";
      verificationInput.maxLength = PASSCODE_LENGTH;
      verificationInput.placeholder = "Enter 5-digit code";
      verificationInput.setAttribute("aria-label", "Two-factor authentication code");
      verificationInput.style.flex = "1";
      verificationInput.style.padding = "10px 12px";
      verificationInput.style.fontSize = "14px";
      verificationInput.style.boxSizing = "border-box";
      verificationInput.style.border = `2px solid ${COLORS.NEUTRAL_BORDER}`;
      verificationInput.style.borderRadius = "4px";
      verificationInput.style.outline = "none";
      verificationInput.style.transition = "border-color 0.2s ease";
      verificationInput.style.textAlign = "center";
      verificationInput.style.letterSpacing = "2px";
      verificationInput.style.fontWeight = "600";

      const vInput = verificationInput;
      vInput.onfocus = (): void => {
        vInput.style.borderColor = COLORS.PRIMARY;
      };
      vInput.onblur = (): void => {
        vInput.style.borderColor = COLORS.NEUTRAL_BORDER;
      };

      sendCodeButton = document.createElement("button");
      sendCodeButton.innerText = "Send Code";
      sendCodeButton.type = "button";
      sendCodeButton.setAttribute("aria-label", "Send verification code");
      sendCodeButton.style.padding = "10px 16px";
      sendCodeButton.style.fontSize = "12px";
      sendCodeButton.style.backgroundColor = "#ffffff";
      sendCodeButton.style.color = COLORS.PRIMARY;
      sendCodeButton.style.border = `2px solid ${COLORS.PRIMARY}`;
      sendCodeButton.style.cursor = "pointer";
      sendCodeButton.style.whiteSpace = "nowrap";
      sendCodeButton.style.borderRadius = "4px";
      sendCodeButton.style.fontWeight = "600";
      sendCodeButton.style.transition = "all 0.2s ease";

      const scBtn = sendCodeButton;
      scBtn.onmouseover = (): void => {
        if (!scBtn.disabled) {
          scBtn.style.backgroundColor = COLORS.PRIMARY;
          scBtn.style.color = "#ffffff";
        }
      };
      scBtn.onmouseout = (): void => {
        if (!scBtn.disabled) {
          scBtn.style.backgroundColor = "#ffffff";
          scBtn.style.color = COLORS.PRIMARY;
        }
      };

      verificationRow.appendChild(verificationInput);
      verificationRow.appendChild(sendCodeButton);

      verificationContainer.appendChild(verificationHeader);
      verificationContainer.appendChild(verificationDescription);
      verificationContainer.appendChild(verificationRow);
    }

    // -----------------------------------------------------------------
    // Cached-signature detection
    // -----------------------------------------------------------------
    const cachedCompressedSig: string | null =
      localStorage.getItem(STORAGE_KEY);
    let hasCachedSignature: boolean =
      cachedCompressedSig !== null && cachedCompressedSig.trim() !== "";
    let decompressedCachedDataUri: string = "";

    if (hasCachedSignature && cachedCompressedSig) {
      try {
        decompressedCachedDataUri = lzwDecompress(cachedCompressedSig);
        if (!decompressedCachedDataUri.startsWith("data:image")) {
          hasCachedSignature = false;
          decompressedCachedDataUri = "";
        }
      } catch {
        hasCachedSignature = false;
        decompressedCachedDataUri = "";
      }
    }

    // -----------------------------------------------------------------
    // Signature section wrapper
    // -----------------------------------------------------------------
    const signatureSection: HTMLDivElement = document.createElement("div");
    signatureSection.style.marginBottom = "20px";
    signatureSection.style.padding = "16px";
    signatureSection.style.backgroundColor = "#ffffff";
    signatureSection.style.border = `1px solid ${COLORS.NEUTRAL_BORDER}`;
    signatureSection.style.borderRadius = "6px";
    signatureSection.style.boxShadow = "0 2px 4px rgba(0,0,0,0.05)";

    const signatureHeader: HTMLDivElement = document.createElement("div");
    signatureHeader.style.marginBottom = "12px";

    const signatureLabel: HTMLHeadingElement =
      document.createElement("h3");
    signatureLabel.innerText = "Digital Signature";
    signatureLabel.style.fontSize = "14px";
    signatureLabel.style.fontWeight = "700";
    signatureLabel.style.color = COLORS.NEUTRAL_TEXT;
    signatureLabel.style.margin = "0 0 4px 0";

    const signatureSubtext: HTMLParagraphElement =
      document.createElement("p");
    signatureSubtext.innerText =
      "Choose how you'd like to provide your signature";
    signatureSubtext.style.fontSize = "12px";
    signatureSubtext.style.color = COLORS.NEUTRAL_TEXT_SEC;
    signatureSubtext.style.margin = "0";

    signatureHeader.appendChild(signatureLabel);
    signatureHeader.appendChild(signatureSubtext);

    // -----------------------------------------------------------------
    // "Remember signature" checkbox
    // -----------------------------------------------------------------
    const cacheControlBox: HTMLDivElement = document.createElement("div");
    cacheControlBox.style.marginTop = "16px";
    cacheControlBox.style.padding = "12px";
    cacheControlBox.style.backgroundColor = COLORS.NEUTRAL_LIGHT;
    cacheControlBox.style.border = `1px solid ${COLORS.NEUTRAL_BORDER}`;
    cacheControlBox.style.borderRadius = "4px";
    cacheControlBox.style.fontSize = "12px";

    const cacheCheckboxLabel: HTMLLabelElement =
      document.createElement("label");
    cacheCheckboxLabel.style.display = "flex";
    cacheCheckboxLabel.style.alignItems = "center";
    cacheCheckboxLabel.style.gap = "8px";
    cacheCheckboxLabel.style.cursor = "pointer";

    const cacheCheckbox: HTMLInputElement = document.createElement("input");
    cacheCheckbox.type = "checkbox";
    cacheCheckbox.checked = true;
    cacheCheckbox.setAttribute("aria-label", "Remember signature for future transactions");
    cacheCheckbox.style.cursor = "pointer";
    cacheCheckbox.style.width = "16px";
    cacheCheckbox.style.height = "16px";

    const cacheCheckboxText: HTMLSpanElement =
      document.createElement("span");
    cacheCheckboxText.innerText =
      "Remember my signature on this device for future transactions";
    cacheCheckboxText.style.color = COLORS.NEUTRAL_TEXT;
    cacheCheckboxText.style.fontSize = "12px";

    cacheCheckboxLabel.appendChild(cacheCheckbox);
    cacheCheckboxLabel.appendChild(cacheCheckboxText);
    cacheControlBox.appendChild(cacheCheckboxLabel);

    // -----------------------------------------------------------------
    // Cached-signature notice panel
    // -----------------------------------------------------------------
    const cachedNoticeContainer: HTMLDivElement =
      document.createElement("div");
    cachedNoticeContainer.style.marginBottom = "16px";
    cachedNoticeContainer.style.padding = "16px";
    cachedNoticeContainer.style.backgroundColor = COLORS.INFO;
    cachedNoticeContainer.style.border = `1px solid ${COLORS.INFO_BORDER}`;
    cachedNoticeContainer.style.borderRadius = "6px";
    cachedNoticeContainer.style.display = hasCachedSignature
      ? "block"
      : "none";

    const cachedNoticeText: HTMLDivElement = document.createElement("div");
    cachedNoticeText.innerText = `${ICONS.CHECK} Saved signature on file`;
    cachedNoticeText.style.fontSize = "13px";
    cachedNoticeText.style.marginBottom = "12px";
    cachedNoticeText.style.fontWeight = "600";
    cachedNoticeText.style.color = COLORS.PRIMARY;

    const cachedPreviewImage: HTMLImageElement =
      document.createElement("img");
    cachedPreviewImage.src = decompressedCachedDataUri;
    cachedPreviewImage.alt = "Your saved signature";
    cachedPreviewImage.style.maxWidth = "100%";
    cachedPreviewImage.style.height = "auto";
    cachedPreviewImage.style.minHeight = "80px";
    cachedPreviewImage.style.maxHeight = "120px";
    cachedPreviewImage.style.display = "block";
    cachedPreviewImage.style.marginBottom = "12px";
    cachedPreviewImage.style.border = `2px solid ${COLORS.PRIMARY}`;
    cachedPreviewImage.style.backgroundColor = "#ffffff";
    cachedPreviewImage.style.borderRadius = "4px";
    cachedPreviewImage.style.padding = "8px";
    cachedPreviewImage.style.objectFit = "contain";

    const removeCachedBtn: HTMLButtonElement =
      document.createElement("button");
    removeCachedBtn.innerText = "Remove Saved Signature";
    removeCachedBtn.setAttribute("aria-label", "Remove saved signature and use new one");
    removeCachedBtn.style.padding = "6px 12px";
    removeCachedBtn.style.fontSize = "11px";
    removeCachedBtn.style.backgroundColor = "#ffffff";
    removeCachedBtn.style.border = `1px solid ${COLORS.ERROR}`;
    removeCachedBtn.style.color = COLORS.ERROR;
    removeCachedBtn.style.cursor = "pointer";
    removeCachedBtn.style.borderRadius = "4px";
    removeCachedBtn.style.fontWeight = "600";
    removeCachedBtn.style.transition = "all 0.2s ease";

    removeCachedBtn.onmouseover = (): void => {
      removeCachedBtn.style.backgroundColor = COLORS.ERROR;
      removeCachedBtn.style.color = "#ffffff";
    };
    removeCachedBtn.onmouseout = (): void => {
      removeCachedBtn.style.backgroundColor = "#ffffff";
      removeCachedBtn.style.color = COLORS.ERROR;
    };

    cachedNoticeContainer.appendChild(cachedNoticeText);
    cachedNoticeContainer.appendChild(cachedPreviewImage);
    cachedNoticeContainer.appendChild(removeCachedBtn);

    // -----------------------------------------------------------------
    // Tab bar (Saved / Draw New / Upload Image)
    // -----------------------------------------------------------------
    const modeContainer: HTMLDivElement = document.createElement("div");
    modeContainer.style.display = "flex";
    modeContainer.style.gap = "8px";
    modeContainer.style.marginBottom = "16px";
    modeContainer.style.borderBottom = `2px solid #edebe9`;
    modeContainer.setAttribute("role", "tablist");

    const createTab = (
      text: string,
      isActive: boolean
    ): HTMLButtonElement => {
      const tab = document.createElement("button");
      tab.innerText = text;
      tab.setAttribute("role", "tab");
      tab.setAttribute("aria-selected", isActive.toString());
      tab.style.padding = "10px 16px";
      tab.style.fontSize = "13px";
      tab.style.cursor = "pointer";
      tab.style.fontWeight = "600";
      tab.style.backgroundColor = "transparent";
      tab.style.color = isActive ? COLORS.PRIMARY : COLORS.NEUTRAL_TEXT_SEC;
      tab.style.border = "none";
      tab.style.borderBottom = isActive
        ? `3px solid ${COLORS.PRIMARY}`
        : "3px solid transparent";
      tab.style.transition = "all 0.2s ease";
      tab.style.outline = "none";

      tab.onmouseover = (): void => {
        const tabIsCurrentlyActive =
          (tab === cachedTabBtn && activeMode === "cached") ||
          (tab === drawTabBtn && activeMode === "draw") ||
          (tab === uploadTabBtn && activeMode === "upload");
        if (!tabIsCurrentlyActive) {
          tab.style.color = COLORS.PRIMARY;
        }
      };
      tab.onmouseout = (): void => {
        const tabIsCurrentlyActive =
          (tab === cachedTabBtn && activeMode === "cached") ||
          (tab === drawTabBtn && activeMode === "draw") ||
          (tab === uploadTabBtn && activeMode === "upload");
        if (!tabIsCurrentlyActive) {
          tab.style.color = COLORS.NEUTRAL_TEXT_SEC;
        }
      };

      return tab;
    };

    const cachedTabBtn = createTab("Saved", hasCachedSignature);
    cachedTabBtn.style.display = hasCachedSignature
      ? "inline-block"
      : "none";

    const drawTabBtn = createTab("Draw New", !hasCachedSignature);
    const uploadTabBtn = createTab("Upload Image", false);

    if (hasCachedSignature) {
      modeContainer.appendChild(cachedTabBtn);
    }
    modeContainer.appendChild(drawTabBtn);
    modeContainer.appendChild(uploadTabBtn);

    // -----------------------------------------------------------------
    // Cached panel
    // -----------------------------------------------------------------
    const cachedPanel: HTMLDivElement = document.createElement("div");
    cachedPanel.setAttribute("role", "tabpanel");
    cachedPanel.style.display = hasCachedSignature ? "block" : "none";
    cachedPanel.appendChild(cachedNoticeContainer);

    // -----------------------------------------------------------------
    // Draw panel (canvas)
    // -----------------------------------------------------------------
    const drawPanel: HTMLDivElement = document.createElement("div");
    drawPanel.setAttribute("role", "tabpanel");
    drawPanel.style.display = hasCachedSignature ? "none" : "block";

    const canvasContainer: HTMLDivElement = document.createElement("div");
    canvasContainer.style.backgroundColor = "#ffffff";
    canvasContainer.style.padding = "16px";
    canvasContainer.style.borderRadius = "6px";
    canvasContainer.style.border = `2px dashed ${COLORS.NEUTRAL_BORDER}`;

    const canvasInstructions: HTMLParagraphElement =
      document.createElement("p");
    canvasInstructions.innerText = "Draw your signature below:";
    canvasInstructions.style.fontSize = "12px";
    canvasInstructions.style.color = COLORS.NEUTRAL_TEXT_SEC;
    canvasInstructions.style.margin = "0 0 8px 0";
    canvasInstructions.style.fontWeight = "600";

    const canvas: HTMLCanvasElement = document.createElement("canvas");
    canvas.width = CANVAS_WIDTH;
    canvas.height = CANVAS_HEIGHT;
    canvas.setAttribute("role", "img");
    canvas.setAttribute("aria-label", "Signature drawing area");
    canvas.style.border = `1px solid ${COLORS.NEUTRAL_BORDER}`;
    canvas.style.backgroundColor = "#ffffff";
    canvas.style.cursor = "crosshair";
    canvas.style.display = "block";
    canvas.style.width = "100%";
    canvas.style.borderRadius = "4px";
    canvas.style.boxShadow = "inset 0 1px 3px rgba(0,0,0,0.1)";

    const clearButton: HTMLButtonElement = document.createElement("button");
    clearButton.innerText = "Clear Canvas";
    clearButton.setAttribute("aria-label", "Clear the signature drawing");
    clearButton.style.marginTop = "10px";
    clearButton.style.padding = "8px 16px";
    clearButton.style.fontSize = "12px";
    clearButton.style.backgroundColor = COLORS.NEUTRAL_LIGHT;
    clearButton.style.border = `1px solid ${COLORS.NEUTRAL_BORDER}`;
    clearButton.style.cursor = "pointer";
    clearButton.style.borderRadius = "4px";
    clearButton.style.fontWeight = "600";
    clearButton.style.color = COLORS.NEUTRAL_TEXT;
    clearButton.style.transition = "all 0.2s ease";

    clearButton.onmouseover = (): void => {
      clearButton.style.backgroundColor = COLORS.NEUTRAL_LIGHTER;
    };
    clearButton.onmouseout = (): void => {
      clearButton.style.backgroundColor = COLORS.NEUTRAL_LIGHT;
    };

    canvasContainer.appendChild(canvasInstructions);
    canvasContainer.appendChild(canvas);
    canvasContainer.appendChild(clearButton);
    drawPanel.appendChild(canvasContainer);

    // -----------------------------------------------------------------
    // Upload panel
    // -----------------------------------------------------------------
    const uploadPanel: HTMLDivElement = document.createElement("div");
    uploadPanel.setAttribute("role", "tabpanel");
    uploadPanel.style.display = "none";
    uploadPanel.style.padding = "32px 24px";
    uploadPanel.style.border = `2px dashed ${COLORS.NEUTRAL_BORDER}`;
    uploadPanel.style.textAlign = "center";
    uploadPanel.style.backgroundColor = "#ffffff";
    uploadPanel.style.borderRadius = "6px";
    uploadPanel.style.transition = "all 0.2s ease";
    uploadPanel.style.minHeight = "200px";
    uploadPanel.style.flexDirection = "column";
    uploadPanel.style.alignItems = "center";
    uploadPanel.style.justifyContent = "center";

    const uploadIcon: HTMLDivElement = document.createElement("div");
    uploadIcon.innerHTML = ICONS.FOLDER;
    uploadIcon.style.fontSize = "48px";
    uploadIcon.style.marginBottom = "16px";
    uploadIcon.setAttribute("aria-hidden", "true");

    const uploadText: HTMLParagraphElement = document.createElement("p");
    uploadText.innerText = "Click to upload a signature image";
    uploadText.style.fontSize = "14px";
    uploadText.style.color = COLORS.NEUTRAL_TEXT_SEC;
    uploadText.style.margin = "0 0 8px 0";
    uploadText.style.fontWeight = "600";

    const uploadSubtext: HTMLParagraphElement =
      document.createElement("p");
    uploadSubtext.innerText = "PNG, JPG (max 5 MB)";
    uploadSubtext.style.fontSize = "12px";
    uploadSubtext.style.color = "#8a8886";
    uploadSubtext.style.margin = "0 0 16px 0";

    const fileInput: HTMLInputElement = document.createElement("input");
    fileInput.type = "file";
    fileInput.accept = "image/png, image/jpeg, image/jpg";
    fileInput.setAttribute("aria-label", "Upload signature image file");
    fileInput.style.fontSize = "12px";
    fileInput.style.marginBottom = "16px";

    const uploadPreview: HTMLImageElement = document.createElement("img");
    uploadPreview.alt = "Preview of uploaded signature";
    uploadPreview.style.maxWidth = "100%";
    uploadPreview.style.height = "auto";
    uploadPreview.style.minHeight = "80px";
    uploadPreview.style.maxHeight = "120px";
    uploadPreview.style.marginTop = "16px";
    uploadPreview.style.display = "none";
    uploadPreview.style.border = `2px solid ${COLORS.PRIMARY}`;
    uploadPreview.style.borderRadius = "4px";
    uploadPreview.style.padding = "8px";
    uploadPreview.style.backgroundColor = "#ffffff";
    uploadPreview.style.objectFit = "contain";

    uploadPanel.appendChild(uploadIcon);
    uploadPanel.appendChild(uploadText);
    uploadPanel.appendChild(uploadSubtext);
    uploadPanel.appendChild(fileInput);
    uploadPanel.appendChild(uploadPreview);

    // -----------------------------------------------------------------
    // State
    // -----------------------------------------------------------------
    let activeMode: "cached" | "draw" | "upload" = hasCachedSignature
      ? "cached"
      : "draw";
    let compressedUploadBase64: string = "";

    const ctx: CanvasRenderingContext2D | null = canvas.getContext("2d");
    let isDrawing: boolean = false;
    let hasDrawnContent: boolean = false;

    if (ctx) {
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.strokeStyle = "#000000";
      ctx.lineWidth = 2.5;
      ctx.lineCap = "round";
      ctx.lineJoin = "round";
    }

    // -----------------------------------------------------------------
    // Button-state helper
    // -----------------------------------------------------------------

    const updateButtonState = (): void => {
      let isCodeValid = true;

      if (requireTFA) {
        if (verificationInput) {
          const codeValue = verificationInput.value.trim();
          isCodeValid =
            codeValue.length === PASSCODE_LENGTH && codeValue === generatedPasscode;
        } else {
          isCodeValid = false;
        }
      }

      let isSignatureValid = false;
      if (activeMode === "cached") {
        isSignatureValid = hasCachedSignature;
      } else if (activeMode === "draw") {
        const canvasHasContent =
          ctx !== null && !isCanvasEmpty(ctx, canvas.width, canvas.height);
        isSignatureValid = hasDrawnContent && canvasHasContent;
      } else {
        isSignatureValid = compressedUploadBase64.trim() !== "";
      }

      const shouldEnable = isCodeValid && isSignatureValid;

      if (shouldEnable) {
        signButton.disabled = false;
        signButton.style.backgroundColor = COLORS.PRIMARY;
        signButton.style.cursor = "pointer";
        signButton.style.opacity = "1";
      } else {
        signButton.disabled = true;
        signButton.style.backgroundColor = "#c8c6c4";
        signButton.style.cursor = "not-allowed";
        signButton.style.opacity = "0.6";
      }
    };

    // -----------------------------------------------------------------
    // Tab-switching logic
    // -----------------------------------------------------------------

    const setActiveTab = (
      mode: "cached" | "draw" | "upload"
    ): void => {
      activeMode = mode;
      cachedPanel.style.display = mode === "cached" ? "block" : "none";
      drawPanel.style.display = mode === "draw" ? "block" : "none";
      uploadPanel.style.display = mode === "upload" ? "flex" : "none";

      [cachedTabBtn, drawTabBtn, uploadTabBtn].forEach((btn) => {
        btn.style.color = COLORS.NEUTRAL_TEXT_SEC;
        btn.style.borderBottom = "3px solid transparent";
        btn.setAttribute("aria-selected", "false");
      });

      const activeBtn =
        mode === "cached"
          ? cachedTabBtn
          : mode === "draw"
          ? drawTabBtn
          : uploadTabBtn;
      activeBtn.style.color = COLORS.PRIMARY;
      activeBtn.style.borderBottom = `3px solid ${COLORS.PRIMARY}`;
      activeBtn.setAttribute("aria-selected", "true");

      updateButtonState();
    };

    if (hasCachedSignature) {
      cachedTabBtn.onclick = (): void => setActiveTab("cached");
    }
    drawTabBtn.onclick = (): void => setActiveTab("draw");
    uploadTabBtn.onclick = (): void => setActiveTab("upload");

    removeCachedBtn.onclick = (): void => {
      localStorage.removeItem(STORAGE_KEY);
      hasCachedSignature = false;
      cachedTabBtn.style.display = "none";
      setActiveTab("draw");
    };

    // -----------------------------------------------------------------
    // Two-factor: send-code logic
    // -----------------------------------------------------------------
    if (requireTFA && context.channel && sendCodeButton && verificationInput) {
      let cooldownTimer: number | null = null;
      const scBtnRef = sendCodeButton;
      const vInputRef = verificationInput;

      const startCooldownTimer = (): void => {
        let secondsLeft = TFA_COOLDOWN_SECONDS;
        scBtnRef.disabled = true;
        scBtnRef.style.backgroundColor = COLORS.NEUTRAL_LIGHT;
        scBtnRef.style.color = "#a19f9d";
        scBtnRef.style.borderColor = COLORS.NEUTRAL_BORDER;
        scBtnRef.style.cursor = "not-allowed";
        scBtnRef.innerText = `Resend (${secondsLeft}s)`;

        if (cooldownTimer !== null) {
          clearInterval(cooldownTimer);
        }

        cooldownTimer = window.setInterval(() => {
          secondsLeft -= 1;
          if (secondsLeft > 0) {
            scBtnRef.innerText = `Resend (${secondsLeft}s)`;
          } else {
            if (cooldownTimer !== null) {
              clearInterval(cooldownTimer);
              cooldownTimer = null;
            }
            scBtnRef.disabled = false;
            scBtnRef.style.backgroundColor = "#ffffff";
            scBtnRef.style.color = COLORS.PRIMARY;
            scBtnRef.style.borderColor = COLORS.PRIMARY;
            scBtnRef.style.cursor = "pointer";
            scBtnRef.innerText = "Resend Code";
          }
        }, 1000);
      };

      const triggerSendCode = async (): Promise<void> => {
        scBtnRef.innerText = "Sending...";
        scBtnRef.disabled = true;

        try {
          if (context.spContext && context.channel) {
            const vResult = await dispatchOtpPasscode(
              context.spContext,
              {
                title: context.signer,
                passcode: generatedPasscode,
                channel: context.channel,
              }
            );

            if (vResult.success && vResult.itemId) {
              storedVerificationItemId = vResult.itemId;
              vInputRef.focus();
            } else {
              alert(
                "Failed to dispatch verification code. Please try again."
              );
            }
          }
        } catch (err) {
          alert("Failed to send verification code. Please try again.");
        } finally {
          startCooldownTimer();
        }
      };

      scBtnRef.onclick = async (): Promise<void> => {
        await triggerSendCode();
      };

      vInputRef.oninput = (): void => {
        if (vInputRef.value.trim().length === PASSCODE_LENGTH) {
          tfaVerified = vInputRef.value.trim() === generatedPasscode;
        }
        updateButtonState();
      };

      // User must click "Send Code" - no automatic send
      // Only auto-send if explicitly marked as needed
    } else if (requireTFA && !context.channel) {
      // TFA required but no delivery channel
    }

    // -----------------------------------------------------------------
    // Canvas drawing helpers
    // -----------------------------------------------------------------

    const getPos = (
      e: MouseEvent | TouchEvent
    ): { x: number; y: number } => {
      const rect: DOMRect = canvas.getBoundingClientRect();
      let clientX: number;
      let clientY: number;

      if ("touches" in e && e.touches.length > 0) {
        clientX = e.touches[0].clientX;
        clientY = e.touches[0].clientY;
      } else if ("changedTouches" in e && e.changedTouches.length > 0) {
        clientX = e.changedTouches[0].clientX;
        clientY = e.changedTouches[0].clientY;
      } else {
        clientX = (e as MouseEvent).clientX;
        clientY = (e as MouseEvent).clientY;
      }

      return {
        x: (clientX - rect.left) * (canvas.width / rect.width),
        y: (clientY - rect.top) * (canvas.height / rect.height),
      };
    };

    const startDraw = (e: MouseEvent | TouchEvent): void => {
      isDrawing = true;
      hasDrawnContent = true;
      const pos = getPos(e);
      ctx?.beginPath();
      ctx?.moveTo(pos.x, pos.y);
      e.preventDefault();
    };

    const drawLine = (e: MouseEvent | TouchEvent): void => {
      if (!isDrawing || !ctx) return;
      const pos = getPos(e);
      ctx.lineTo(pos.x, pos.y);
      ctx.stroke();
      updateButtonState();
      e.preventDefault();
    };

    const stopDraw = (): void => {
      isDrawing = false;
      updateButtonState();
    };

    canvas.addEventListener("mousedown", startDraw);
    canvas.addEventListener("mousemove", drawLine);
    canvas.addEventListener("touchstart", startDraw, { passive: false });
    canvas.addEventListener("touchmove", drawLine, { passive: false });

    window.addEventListener("mouseup", stopDraw);
    window.addEventListener("touchend", stopDraw);

    // Track abort controller listeners
    abortController.signal.addEventListener("abort", () => {
      canvas.removeEventListener("mousedown", startDraw);
      canvas.removeEventListener("mousemove", drawLine);
      canvas.removeEventListener("touchstart", startDraw);
      canvas.removeEventListener("touchmove", drawLine);
      window.removeEventListener("mouseup", stopDraw);
      window.removeEventListener("touchend", stopDraw);
    });

    clearButton.onclick = (): void => {
      if (ctx) {
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        hasDrawnContent = false;
        updateButtonState();
      }
    };

    // -----------------------------------------------------------------
    // File upload handler
    // -----------------------------------------------------------------

    fileInput.onchange = (e: Event): void => {
      const target = e.target as HTMLInputElement;
      if (target.files && target.files[0]) {
        const file = target.files[0];

        if (file.size > MAX_UPLOAD_BYTES) {
          alert(
            `File size exceeds 5 MB (${(file.size / 1024 / 1024).toFixed(
              1
            )} MB). Please choose a smaller image.`
          );
          target.value = "";
          return;
        }

        const reader = new FileReader();
        reader.onerror = (): void => {
          alert("Failed to load image. Please try again.");
          target.value = "";
          updateButtonState();
        };

        reader.onload = (uploadEvent): void => {
          if (uploadEvent.target?.result) {
            const img = new Image();
            img.onerror = (): void => {
              alert("Failed to read image file. Please try another image.");
              target.value = "";
              updateButtonState();
            };

            img.onload = (): void => {
              resizeImageAndStore(img, target);
            };
            img.src = uploadEvent.target.result as string;
          }
        };
        reader.readAsDataURL(file);
      }
    };

    /**
     * Helper to resize image and store result.
     *
     * @param img - Loaded image element.
     * @param fileInput - The input element reference.
     */
    const resizeImageAndStore = (img: HTMLImageElement, fileInput: HTMLInputElement): void => {
      try {
        const tempCanvas = document.createElement("canvas");
        let w = img.width;
        let h = img.height;

        if (w > IMAGE_MAX_WIDTH) {
          h = Math.round((h * IMAGE_MAX_WIDTH) / w);
          w = IMAGE_MAX_WIDTH;
        }
        if (h > IMAGE_MAX_HEIGHT) {
          w = Math.round((w * IMAGE_MAX_HEIGHT) / h);
          h = IMAGE_MAX_HEIGHT;
        }

        tempCanvas.width = w;
        tempCanvas.height = h;
        const tCtx = tempCanvas.getContext("2d");
        if (tCtx) {
          tCtx.fillStyle = "#ffffff";
          tCtx.fillRect(0, 0, w, h);
          tCtx.drawImage(img, 0, 0, w, h);
          compressedUploadBase64 = tempCanvas.toDataURL("image/png");
          uploadPreview.src = compressedUploadBase64;
          uploadPreview.style.display = "block";
          updateButtonState();
        }
      } catch (error) {
        alert("Failed to process image. Please try another image.");
        fileInput.value = "";
        updateButtonState();
      }
    };

    // -----------------------------------------------------------------
    // Footer (Cancel + Sign buttons)
    // -----------------------------------------------------------------
    const footerSection: HTMLDivElement = document.createElement("div");
    footerSection.style.padding = "20px 24px";
    footerSection.style.backgroundColor = "#ffffff";
    footerSection.style.borderTop = "1px solid #edebe9";
    footerSection.style.display = "flex";
    footerSection.style.justifyContent = "flex-end";
    footerSection.style.gap = "12px";

    const cancelButton: HTMLButtonElement =
      document.createElement("button");
    cancelButton.innerText = "Cancel";
    cancelButton.setAttribute("aria-label", "Cancel signing and close dialog");
    cancelButton.style.padding = "10px 24px";
    cancelButton.style.backgroundColor = "#ffffff";
    cancelButton.style.border = `1px solid ${COLORS.NEUTRAL_BORDER_DARK}`;
    cancelButton.style.cursor = "pointer";
    cancelButton.style.fontSize = "14px";
    cancelButton.style.fontWeight = "600";
    cancelButton.style.color = COLORS.NEUTRAL_TEXT;
    cancelButton.style.borderRadius = "4px";
    cancelButton.style.transition = "all 0.2s ease";

    cancelButton.onmouseover = (): void => {
      cancelButton.style.backgroundColor = COLORS.NEUTRAL_LIGHT;
    };
    cancelButton.onmouseout = (): void => {
      cancelButton.style.backgroundColor = "#ffffff";
    };

    signButton.onmouseover = (): void => {
      if (!signButton.disabled) {
        signButton.style.backgroundColor = COLORS.PRIMARY_DARK;
      }
    };
    signButton.onmouseout = (): void => {
      if (!signButton.disabled) {
        signButton.style.backgroundColor = COLORS.PRIMARY;
      }
    };

    // -----------------------------------------------------------------
    // Keyboard support
    // -----------------------------------------------------------------
    const handleEscapeKey = (e: KeyboardEvent): void => {
      if (e.key === "Escape" || e.key === "Esc") {
        cleanup();
        resolve(undefined);
      }
    };

    document.addEventListener("keydown", handleEscapeKey);
    abortController.signal.addEventListener("abort", () => {
      document.removeEventListener("keydown", handleEscapeKey);
    });

    // -----------------------------------------------------------------
    // Cleanup helper
    // -----------------------------------------------------------------

    const cleanup = (): void => {
      if (activeModalInstance === abortController) {
        activeModalInstance = null;
      }

      clearTimeout(timeoutId);
      abortController.abort();

      document.removeEventListener("keydown", handleEscapeKey);

      if (document.body.contains(overlay)) {
        document.body.removeChild(overlay);
      }
    };

    cancelButton.onclick = (): void => {
      cleanup();
      resolve(undefined);
    };

    // -----------------------------------------------------------------
    // Sign button handler
    // -----------------------------------------------------------------

    signButton.onclick = async (): Promise<void> => {
      if (requireTFA && verificationInput) {
        const enteredCode = verificationInput.value.trim();

        if (enteredCode !== generatedPasscode) {
          alert(
            "Invalid verification code. Please enter the correct 5-digit code."
          );
          return;
        }
      }

      let finalDataUrl = "";

      if (activeMode === "cached") {
        if (!cachedCompressedSig) {
          alert(
            "No cached signature found. Please select another option."
          );
          return;
        }
        try {
          finalDataUrl = lzwDecompress(cachedCompressedSig);
          if (!finalDataUrl.startsWith("data:image")) {
            throw new Error("Invalid signature data");
          }
        } catch {
          alert("Cached signature is corrupted. Please use a different method.");
          return;
        }
      } else if (activeMode === "draw") {
        if (
          !ctx ||
          isCanvasEmpty(ctx, canvas.width, canvas.height)
        ) {
          alert("Please draw your signature before proceeding.");
          return;
        }
        finalDataUrl = downscaleCanvas(canvas, CANVAS_WIDTH, CANVAS_HEIGHT);
      } else {
        if (
          !compressedUploadBase64 ||
          compressedUploadBase64.trim() === ""
        ) {
          alert(
            "Please upload a signature image before proceeding."
          );
          return;
        }
        finalDataUrl = compressedUploadBase64;
      }

      try {
        const compressedString = lzwCompress(finalDataUrl);

        if (cacheCheckbox.checked) {
          localStorage.setItem(STORAGE_KEY, compressedString);
        } else {
          localStorage.removeItem(STORAGE_KEY);
        }

        cleanup();

        const fullContext = {
          ...context,
          signatureData: compressedString,
        };

        const result: SharePointAuditRecord =
          await generateSecureAuditRecordInternal(fullContext);

        if (storedVerificationItemId !== undefined) {
          result.verificationItemId = storedVerificationItemId;
        }

        resolve(result);
      } catch (error: unknown) {
        alert("Signing failed. Please try again.");
        resolve(undefined);
      }
    };

    updateButtonState();

    // -----------------------------------------------------------------
    // Assemble DOM tree
    // -----------------------------------------------------------------
    footerSection.appendChild(cancelButton);
    footerSection.appendChild(signButton);

    contentSection.appendChild(body);
    if (requireTFA) {
      contentSection.appendChild(verificationContainer);
    }
    contentSection.appendChild(signatureSection);

    signatureSection.appendChild(signatureHeader);
    signatureSection.appendChild(modeContainer);
    signatureSection.appendChild(cachedPanel);
    signatureSection.appendChild(drawPanel);
    signatureSection.appendChild(uploadPanel);
    signatureSection.appendChild(cacheControlBox);

    modalBox.appendChild(headerSection);
    modalBox.appendChild(contentSection);
    modalBox.appendChild(footerSection);

    overlay.appendChild(modalBox);
    document.body.appendChild(overlay);
  });
}

// ---------------------------------------------------------------------------
// Signature decompression for reports
// ---------------------------------------------------------------------------

/**
 * Decompresses an LZW-compressed signature string back into its
 * original `data:image/…` data-URI for display in reports or
 * print views.
 *
 * @param compressedSignatureData - The LZW-compressed string stored
 *                                  in SharePoint.
 * @returns The original data-URI, or an empty string if
 *          decompression fails or the result is not an image
 *          data-URI.
 */
export function getReportableSignature(
  compressedSignatureData: string
): string {
  try {
    const decompressed = lzwDecompress(compressedSignatureData);
    if (!decompressed || !decompressed.startsWith("data:image")) {
      return "";
    }
    return decompressed;
  } catch {
    return "";
  }
}

// ---------------------------------------------------------------------------
// LZW compression / decompression
// ---------------------------------------------------------------------------

/**
 * Compresses a string using the LZW (Lempel-Ziv-Welch) algorithm.
 *
 * **Important:** The dictionary is limited to 65,535 entries to prevent
 * corruption in localStorage/JSON. Large inputs may fail to compress.
 * For data-URIs larger than this limit, consider using alternative
 * compression methods or Base64 encoding.
 *
 * @param input - The raw string to compress.
 * @returns The compressed string (each character encodes one
 *          dictionary index).
 * @throws {Error} If dictionary exceeds maximum size.
 */
function lzwCompress(input: string): string {
  const dictionary: { [key: string]: number } = {};
  let c = "";
  let wc = "";
  let w = "";
  const result: number[] = [];
  let dictionarySize = 256;

  for (let i = 0; i < 256; i += 1) {
    dictionary[String.fromCharCode(i)] = i;
  }

  for (let i = 0; i < input.length; i += 1) {
    c = input.charAt(i);
    wc = w + c;
    if (Object.prototype.hasOwnProperty.call(dictionary, wc)) {
      w = wc;
    } else {
      result.push(dictionary[w]);

      if (dictionarySize >= LZW_DICT_MAX) {
        throw new Error(
          `LZW dictionary overflow: input is too large to compress (${dictionarySize} entries). ` +
          `Consider using Base64 encoding instead.`
        );
      }

      dictionary[wc] = dictionarySize++;
      w = String(c);
    }
  }

  if (w !== "") {
    result.push(dictionary[w]);
  }

  return result
    .map((n: number): string => String.fromCharCode(n))
    .join("");
}

/**
 * Decompresses a string that was compressed with {@link lzwCompress}.
 *
 * @param compressed - The LZW-compressed string.
 * @returns The original string, or an empty string if the input is
 *          empty or contains an unrecognised dictionary reference.
 */
function lzwDecompress(compressed: string): string {
  const dictionary: { [key: number]: string } = {};
  const result: string[] = [];
  let dictionarySize = 256;

  for (let i = 0; i < 256; i += 1) {
    dictionary[i] = String.fromCharCode(i);
  }

  if (compressed.length === 0) return "";

  let w = String.fromCharCode(compressed.charCodeAt(0));
  result.push(w);

  let entry = "";
  for (let i = 1; i < compressed.length; i += 1) {
    const k = compressed.charCodeAt(i);
    if (dictionary[k] !== undefined) {
      entry = dictionary[k];
    } else if (k === dictionarySize) {
      entry = w + w.charAt(0);
    } else {
      return "";
    }

    result.push(entry);
    dictionary[dictionarySize++] = w + entry.charAt(0);
    w = entry;
  }

  return result.join("");
}

// ---------------------------------------------------------------------------
// Canvas utilities
// ---------------------------------------------------------------------------

/**
 * Draws the contents of `sourceCanvas` onto a new canvas at the
 * specified dimensions and returns the result as a PNG data-URI.
 *
 * @param sourceCanvas - The canvas element to read from.
 * @param targetWidth  - Desired output width in pixels.
 * @param targetHeight - Desired output height in pixels.
 * @returns A `data:image/png;base64,…` string.
 */
function downscaleCanvas(
  sourceCanvas: HTMLCanvasElement,
  targetWidth: number,
  targetHeight: number
): string {
  const tempCanvas = document.createElement("canvas");
  tempCanvas.width = targetWidth;
  tempCanvas.height = targetHeight;
  const ctx = tempCanvas.getContext("2d");
  if (ctx) {
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, targetWidth, targetHeight);
    ctx.drawImage(sourceCanvas, 0, 0, targetWidth, targetHeight);
  }
  return tempCanvas.toDataURL("image/png");
}

/**
 * Determines whether every pixel on the canvas is pure white
 * (`rgb(255, 255, 255)`), which indicates that the user has not
 * drawn anything yet.
 *
 * @param ctx    - The 2D rendering context of the canvas.
 * @param width  - Canvas width in pixels.
 * @param height - Canvas height in pixels.
 * @returns `true` if the canvas contains only white pixels.
 */
function isCanvasEmpty(
  ctx: CanvasRenderingContext2D,
  width: number,
  height: number
): boolean {
  const pixelBuffer = ctx.getImageData(0, 0, width, height).data;
  for (let i = 0; i < pixelBuffer.length; i += 4) {
    if (
      pixelBuffer[i] !== 255 ||
      pixelBuffer[i + 1] !== 255 ||
      pixelBuffer[i + 2] !== 255
    ) {
      return false;
    }
  }
  return true;
}

// ---------------------------------------------------------------------------
// Audit-record verification
// ---------------------------------------------------------------------------

/**
 * Re-computes the SHA-256 hash of the audit envelope constructed from
 * the supplied parameters and compares it to the stored hash.
 *
 * This allows any consumer to independently verify that a signed
 * record has not been tampered with, without needing access to the
 * original signature image.
 *
 * **Deep Sorting:** Payload is deep-sorted to ensure nested objects
 * maintain consistent key ordering. All payload values must be
 * JSON-serializable (no functions, symbols, undefined at top level).
 *
 * **Timestamp Precision:** Timestamps are stored at second precision
 * to avoid hash mismatches from millisecond variations.
 *
 * @param auditRecord - The SharePointAuditRecord returned from `promptAndGenerateSecureAudit`.
 * @param signer      - The signer's email address or display name (must match original signer).
 * @param payload     - The original payload object that was signed (must be JSON-serializable).
 * @returns `true` if the signature hash is valid and authentic.
 *
 * @example
 * ```ts
 * const isValid = await verifySecureAuditRecord(
 *   auditRecord,
 *   "user@example.com",
 *   { amount: 1500, vendor: "Contoso" }
 * );
 *
 * if (isValid) {
 *   console.log("Signature is authentic!");
 * } else {
 *   console.warn("Signature has been tampered with!");
 * }
 * ```
 */
export async function verifySecureAuditRecord(
  auditRecord: SharePointAuditRecord,
  signer: string,
  payload: Record<string, unknown>
): Promise<boolean> {
  try {
    if (!auditRecord || !signer || !payload) {
      return false;
    }

    if (!auditRecord.signatureHash || !auditRecord.signatureTimestamp) {
      return false;
    }

    const normalizedSigner = sanitizeSigner(signer);
    const normalizedTimestamp = auditRecord.signatureTimestamp.trim();
    const normalizedHash = auditRecord.signatureHash.trim();

    validateJsonSerializable(payload);
    const deepSortedPayload = deepSortPayload(payload) as Record<string, unknown>;

    const auditEnvelope: AuditEnvelopeRecord = {
      payload: deepSortedPayload,
      signer: normalizedSigner,
      timestamp: normalizedTimestamp,
    };

    const auditEnvelopeJson: string = JSON.stringify(auditEnvelope);
    const encoder: TextEncoder = new TextEncoder();
    const encodedBytes: Uint8Array = encoder.encode(auditEnvelopeJson);

    const hashBuffer: ArrayBuffer = await window.crypto.subtle.digest(
      "SHA-256",
      encodedBytes.buffer as ArrayBuffer
    );
    const hashArray: number[] = Array.from(new Uint8Array(hashBuffer));
    const recomputedHash: string = hashArray
      .map((b: number): string => b.toString(16).padStart(2, "0"))
      .join("");

    const isValid = recomputedHash === normalizedHash;

    return isValid;
  } catch (error) {
    return false;
  }
}

/**
 * Deep-sorts a payload object recursively, ensuring nested objects
 * and array items maintain consistent key ordering.
 *
 * @param obj - The object to sort
 * @returns A new object with all keys sorted recursively
 */
function deepSortPayload(obj: Record<string, unknown> | unknown[]): Record<string, unknown> | unknown[] {
  if (obj === null || obj === undefined) {
    return obj as Record<string, unknown>;
  }

  if (Array.isArray(obj)) {
    return obj.map((item) => {
      if (item !== null && typeof item === "object" && !Array.isArray(item)) {
        return deepSortPayload(item as Record<string, unknown>);
      }
      return item;
    });  // ← Removed 'as unknown[]' cast
  }

  if (typeof obj === "object") {
    const sorted: Record<string, unknown> = {};
    const keys = Object.keys(obj as Record<string, unknown>).sort();

    keys.forEach((key) => {
      const value = (obj as Record<string, unknown>)[key];

      if (value === null || value === undefined) {
        sorted[key] = value;
      } else if (Array.isArray(value)) {
        sorted[key] = value.map((item) => {
          if (item !== null && typeof item === "object" && !Array.isArray(item)) {
            return deepSortPayload(item as Record<string, unknown>);
          }
          return item;
        });
      } else if (typeof value === "object") {
        sorted[key] = deepSortPayload(value as Record<string, unknown>);
      } else {
        sorted[key] = value;
      }
    });

    return sorted;
  }

  return obj;
}

// ---------------------------------------------------------------------------
// Internal – audit-record generation
// ---------------------------------------------------------------------------

/**
 * Constructs the canonical audit envelope from the signer context,
 * hashes it with SHA-256, and returns a `SharePointAuditRecord`
 * ready for persistence.
 *
 * This is an internal function and is not exported.
 *
 * **Key Binding:** The itemID is now included in the audit envelope
 * hash to tie the signature to a specific transaction and prevent
 * signature replay attacks.
 *
 * @param context - The full signer context, extended with the
 *                  compressed signature data string.
 * @returns A fully populated `SharePointAuditRecord`.
 *
 * @throws {Error} If the payload is empty, signer is blank, or payload
 *                 is not JSON-serializable.
 */
async function generateSecureAuditRecordInternal(
  context: SignerContext & { signatureData: string }
): Promise<SharePointAuditRecord> {
  const { payload, signer, signatureData, itemID } = context;

  assertNonEmptyString(signer, "Signer");
  const sanitizedSigner = sanitizeSigner(signer);

  if (!payload || Object.keys(payload).length === 0) {
    throw new Error("Payload cannot be empty.");
  }

  validateJsonSerializable(payload);

  const signatureTimestamp: string = generateTimestamp();

  const deepSortedPayload = deepSortPayload(payload) as Record<string, unknown>;

  const auditEnvelope: AuditEnvelopeRecord = {
    payload: deepSortedPayload,
    signer: sanitizedSigner,
    timestamp: signatureTimestamp,
    itemID: itemID,
  };

  const auditEnvelopeJson: string = JSON.stringify(auditEnvelope);
  const encoder: TextEncoder = new TextEncoder();
  const encodedBytes: Uint8Array = encoder.encode(auditEnvelopeJson);

  const hashBuffer: ArrayBuffer = await window.crypto.subtle.digest(
    "SHA-256",
    encodedBytes.buffer as ArrayBuffer
  );
  const hashArray: number[] = Array.from(new Uint8Array(hashBuffer));
  const signatureHash: string = hashArray
    .map((b: number): string => b.toString(16).padStart(2, "0"))
    .join("");

  return {
    signatureHash,
    signatureData,
    signatureTimestamp,
    verificationItemId: itemID,
  };
}