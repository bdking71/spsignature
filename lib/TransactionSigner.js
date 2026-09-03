"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.verifySecureAuditRecord = exports.getReportableSignature = exports.promptAndGenerateSecureAudit = void 0;
const OtpService_1 = require("./OtpService");
/** localStorage key used to persist a compressed signature across sessions. */
const STORAGE_KEY = "secure_audit_cached_signature_v1";
/** Maximum upload file size in bytes (5 MB). */
const MAX_UPLOAD_BYTES = 5 * 1024 * 1024;
// ---------------------------------------------------------------------------
// Helpers – passcode generation
// ---------------------------------------------------------------------------
/**
 * Generates a cryptographically random 5-digit passcode (10 000 – 99 999)
 * using the Web Crypto API.
 *
 * @returns A string representation of the 5-digit code.
 */
function generateFiveDigitPasscode() {
    const array = new Uint32Array(1);
    window.crypto.getRandomValues(array);
    const code = (array[0] % 90000) + 10000;
    return code.toString();
}
// ---------------------------------------------------------------------------
// Main public function – modal signing ceremony
// ---------------------------------------------------------------------------
/**
 * Opens a full-screen modal dialog that walks the user through:
 *
 * 1. (Optional) Two-factor passcode verification
 * 2. Providing a digital signature (cached / drawn / uploaded)
 * 3. Generating a tamper-evident SHA-256 audit record
 *
 * The returned promise resolves with a `SharePointAuditRecord` on
 * success or `undefined` if the user cancels.
 *
 * @param context        - Signer metadata and SPFx context.
 * @param modalTitle     - Title shown in the modal header.
 * @param warningMessage - Instructional HTML rendered above the
 *                         signature area.
 * @returns A promise that resolves to the audit record or `undefined`.
 *
 * @throws {Error} If `context.spContext` is falsy.
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
 * ```
 */
async function promptAndGenerateSecureAudit(context, modalTitle = "Approve Purchase Requisition", warningMessage = "Please review your entry and apply your signature to finalize this record.") {
    if (!context.spContext) {
        throw new Error("Execution restricted: Valid WebPartContext must be supplied.");
    }
    return new Promise((resolve) => {
        // TFA is only enabled if explicitly set to true
        const requireTFA = context.requireTFA === true;
        const generatedPasscode = generateFiveDigitPasscode();
        let storedVerificationItemId = undefined;
        console.log("TFA Enabled:", requireTFA); // DEBUG
        // -----------------------------------------------------------------
        // Overlay
        // -----------------------------------------------------------------
        const overlay = document.createElement("div");
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
        const modalBox = document.createElement("div");
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
        const headerSection = document.createElement("div");
        headerSection.style.backgroundColor = "#0078d4";
        headerSection.style.padding = "20px 24px";
        headerSection.style.borderBottom = "3px solid #005a9e";
        const title = document.createElement("h2");
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
        const contentSection = document.createElement("div");
        contentSection.style.padding = "24px";
        contentSection.style.backgroundColor = "#fafafa";
        contentSection.style.overflowY = "auto";
        contentSection.style.flex = "1";
        const body = document.createElement("p");
        body.innerHTML = warningMessage;
        body.style.fontSize = "14px";
        body.style.color = "#323130";
        body.style.lineHeight = "1.6";
        body.style.margin = "0 0 20px 0";
        // -----------------------------------------------------------------
        // Sign button (footer)
        // -----------------------------------------------------------------
        const signButton = document.createElement("button");
        signButton.innerText = "I Agree and Sign";
        signButton.style.padding = "10px 24px";
        signButton.style.backgroundColor = "#0078d4";
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
        const verificationContainer = document.createElement("div");
        /** Direct reference to the passcode <input> (avoids fragile querySelector). */
        let verificationInput = null;
        /** Direct reference to the "Resend Code" button. */
        let sendCodeButton = null;
        if (requireTFA) {
            verificationContainer.style.marginBottom = "20px";
            verificationContainer.style.padding = "16px";
            verificationContainer.style.backgroundColor = "#fff8e1";
            verificationContainer.style.border = "1px solid #ffd54f";
            verificationContainer.style.borderRadius = "6px";
            verificationContainer.style.boxShadow =
                "0 2px 4px rgba(0,0,0,0.05)";
            /* Header row with lock icon */
            const verificationHeader = document.createElement("div");
            verificationHeader.style.display = "flex";
            verificationHeader.style.alignItems = "center";
            verificationHeader.style.marginBottom = "12px";
            verificationHeader.style.gap = "8px";
            const lockIcon = document.createElement("span");
            lockIcon.innerHTML = "🔐";
            lockIcon.style.fontSize = "18px";
            const verificationLabel = document.createElement("label");
            verificationLabel.innerText = "Two-Factor Authentication";
            verificationLabel.style.display = "block";
            verificationLabel.style.fontSize = "13px";
            verificationLabel.style.fontWeight = "700";
            verificationLabel.style.color = "#323130";
            verificationLabel.style.margin = "0";
            verificationHeader.appendChild(lockIcon);
            verificationHeader.appendChild(verificationLabel);
            const verificationDescription = document.createElement("p");
            verificationDescription.innerText = `A 5-digit verification code has been sent to ${context.channel === "teams" ? "Microsoft Teams" : "your email"}.`;
            verificationDescription.style.fontSize = "12px";
            verificationDescription.style.color = "#605e5c";
            verificationDescription.style.margin = "0 0 12px 0";
            verificationDescription.style.lineHeight = "1.5";
            /* Input + Resend row */
            const verificationRow = document.createElement("div");
            verificationRow.style.display = "flex";
            verificationRow.style.gap = "10px";
            verificationInput = document.createElement("input");
            verificationInput.type = "text";
            verificationInput.maxLength = 5;
            verificationInput.placeholder = "Enter 5-digit code";
            verificationInput.style.flex = "1";
            verificationInput.style.padding = "10px 12px";
            verificationInput.style.fontSize = "14px";
            verificationInput.style.boxSizing = "border-box";
            verificationInput.style.border = "2px solid #d1d1d1";
            verificationInput.style.borderRadius = "4px";
            verificationInput.style.outline = "none";
            verificationInput.style.transition = "border-color 0.2s ease";
            verificationInput.style.textAlign = "center";
            verificationInput.style.letterSpacing = "2px";
            verificationInput.style.fontWeight = "600";
            const vInput = verificationInput; // capture for closures
            vInput.onfocus = () => {
                vInput.style.borderColor = "#0078d4";
            };
            vInput.onblur = () => {
                vInput.style.borderColor = "#d1d1d1";
            };
            sendCodeButton = document.createElement("button");
            sendCodeButton.innerText = "Resend Code";
            sendCodeButton.type = "button";
            sendCodeButton.style.padding = "10px 16px";
            sendCodeButton.style.fontSize = "12px";
            sendCodeButton.style.backgroundColor = "#ffffff";
            sendCodeButton.style.color = "#0078d4";
            sendCodeButton.style.border = "2px solid #0078d4";
            sendCodeButton.style.cursor = "pointer";
            sendCodeButton.style.whiteSpace = "nowrap";
            sendCodeButton.style.borderRadius = "4px";
            sendCodeButton.style.fontWeight = "600";
            sendCodeButton.style.transition = "all 0.2s ease";
            const scBtn = sendCodeButton; // capture for closures
            scBtn.onmouseover = () => {
                scBtn.style.backgroundColor = "#0078d4";
                scBtn.style.color = "#ffffff";
            };
            scBtn.onmouseout = () => {
                if (!scBtn.disabled) {
                    scBtn.style.backgroundColor = "#ffffff";
                    scBtn.style.color = "#0078d4";
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
        const cachedCompressedSig = localStorage.getItem(STORAGE_KEY);
        let hasCachedSignature = cachedCompressedSig !== null && cachedCompressedSig.trim() !== "";
        const decompressedCachedDataUri = hasCachedSignature && cachedCompressedSig
            ? lzwDecompress(cachedCompressedSig)
            : "";
        // -----------------------------------------------------------------
        // Signature section wrapper
        // -----------------------------------------------------------------
        const signatureSection = document.createElement("div");
        signatureSection.style.marginBottom = "20px";
        signatureSection.style.padding = "16px";
        signatureSection.style.backgroundColor = "#ffffff";
        signatureSection.style.border = "1px solid #d1d1d1";
        signatureSection.style.borderRadius = "6px";
        signatureSection.style.boxShadow = "0 2px 4px rgba(0,0,0,0.05)";
        const signatureHeader = document.createElement("div");
        signatureHeader.style.marginBottom = "12px";
        const signatureLabel = document.createElement("h3");
        signatureLabel.innerText = "Digital Signature";
        signatureLabel.style.fontSize = "14px";
        signatureLabel.style.fontWeight = "700";
        signatureLabel.style.color = "#323130";
        signatureLabel.style.margin = "0 0 4px 0";
        const signatureSubtext = document.createElement("p");
        signatureSubtext.innerText =
            "Choose how you'd like to provide your signature";
        signatureSubtext.style.fontSize = "12px";
        signatureSubtext.style.color = "#605e5c";
        signatureSubtext.style.margin = "0";
        signatureHeader.appendChild(signatureLabel);
        signatureHeader.appendChild(signatureSubtext);
        // -----------------------------------------------------------------
        // "Remember signature" checkbox
        // -----------------------------------------------------------------
        const cacheControlBox = document.createElement("div");
        cacheControlBox.style.marginTop = "16px";
        cacheControlBox.style.padding = "12px";
        cacheControlBox.style.backgroundColor = "#f3f2f1";
        cacheControlBox.style.border = "1px solid #d1d1d1";
        cacheControlBox.style.borderRadius = "4px";
        cacheControlBox.style.fontSize = "12px";
        const cacheCheckboxLabel = document.createElement("label");
        cacheCheckboxLabel.style.display = "flex";
        cacheCheckboxLabel.style.alignItems = "center";
        cacheCheckboxLabel.style.gap = "8px";
        cacheCheckboxLabel.style.cursor = "pointer";
        const cacheCheckbox = document.createElement("input");
        cacheCheckbox.type = "checkbox";
        cacheCheckbox.checked = true;
        cacheCheckbox.style.cursor = "pointer";
        cacheCheckbox.style.width = "16px";
        cacheCheckbox.style.height = "16px";
        const cacheCheckboxText = document.createElement("span");
        cacheCheckboxText.innerText =
            "Remember my signature on this device for future transactions";
        cacheCheckboxText.style.color = "#323130";
        cacheCheckboxText.style.fontSize = "12px";
        cacheCheckboxLabel.appendChild(cacheCheckbox);
        cacheCheckboxLabel.appendChild(cacheCheckboxText);
        cacheControlBox.appendChild(cacheCheckboxLabel);
        // -----------------------------------------------------------------
        // Cached-signature notice panel
        // -----------------------------------------------------------------
        const cachedNoticeContainer = document.createElement("div");
        cachedNoticeContainer.style.marginBottom = "16px";
        cachedNoticeContainer.style.padding = "16px";
        cachedNoticeContainer.style.backgroundColor = "#e6f4ff";
        cachedNoticeContainer.style.border = "1px solid #0078d4";
        cachedNoticeContainer.style.borderRadius = "6px";
        cachedNoticeContainer.style.display = hasCachedSignature
            ? "block"
            : "none";
        const cachedNoticeText = document.createElement("div");
        cachedNoticeText.innerText = "✓ Saved signature on file";
        cachedNoticeText.style.fontSize = "13px";
        cachedNoticeText.style.marginBottom = "12px";
        cachedNoticeText.style.fontWeight = "600";
        cachedNoticeText.style.color = "#0078d4";
        const cachedPreviewImage = document.createElement("img");
        cachedPreviewImage.src = decompressedCachedDataUri;
        cachedPreviewImage.style.maxWidth = "100%";
        cachedPreviewImage.style.height = "auto";
        cachedPreviewImage.style.minHeight = "80px";
        cachedPreviewImage.style.maxHeight = "120px";
        cachedPreviewImage.style.display = "block";
        cachedPreviewImage.style.marginBottom = "12px";
        cachedPreviewImage.style.border = "2px solid #0078d4";
        cachedPreviewImage.style.backgroundColor = "#ffffff";
        cachedPreviewImage.style.borderRadius = "4px";
        cachedPreviewImage.style.padding = "8px";
        cachedPreviewImage.style.objectFit = "contain";
        const removeCachedBtn = document.createElement("button");
        removeCachedBtn.innerText = "Remove Saved Signature";
        removeCachedBtn.style.padding = "6px 12px";
        removeCachedBtn.style.fontSize = "11px";
        removeCachedBtn.style.backgroundColor = "#ffffff";
        removeCachedBtn.style.border = "1px solid #a4262c";
        removeCachedBtn.style.color = "#a4262c";
        removeCachedBtn.style.cursor = "pointer";
        removeCachedBtn.style.borderRadius = "4px";
        removeCachedBtn.style.fontWeight = "600";
        removeCachedBtn.style.transition = "all 0.2s ease";
        removeCachedBtn.onmouseover = () => {
            removeCachedBtn.style.backgroundColor = "#a4262c";
            removeCachedBtn.style.color = "#ffffff";
        };
        removeCachedBtn.onmouseout = () => {
            removeCachedBtn.style.backgroundColor = "#ffffff";
            removeCachedBtn.style.color = "#a4262c";
        };
        cachedNoticeContainer.appendChild(cachedNoticeText);
        cachedNoticeContainer.appendChild(cachedPreviewImage);
        cachedNoticeContainer.appendChild(removeCachedBtn);
        // -----------------------------------------------------------------
        // Tab bar (Saved / Draw New / Upload Image)
        // -----------------------------------------------------------------
        const modeContainer = document.createElement("div");
        modeContainer.style.display = "flex";
        modeContainer.style.gap = "8px";
        modeContainer.style.marginBottom = "16px";
        modeContainer.style.borderBottom = "2px solid #edebe9";
        /**
         * Factory that creates a styled tab button.
         *
         * @param text      - Label for the tab.
         * @param isActive  - Whether the tab should appear selected initially.
         * @returns The constructed `<button>` element.
         */
        const createTab = (text, isActive) => {
            const tab = document.createElement("button");
            tab.innerText = text;
            tab.style.padding = "10px 16px";
            tab.style.fontSize = "13px";
            tab.style.cursor = "pointer";
            tab.style.fontWeight = "600";
            tab.style.backgroundColor = "transparent";
            tab.style.color = isActive ? "#0078d4" : "#605e5c";
            tab.style.border = "none";
            tab.style.borderBottom = isActive
                ? "3px solid #0078d4"
                : "3px solid transparent";
            tab.style.transition = "all 0.2s ease";
            tab.style.outline = "none";
            /* Hover handlers check current activeMode so they remain correct
               after the user switches tabs. */
            tab.onmouseover = () => {
                const tabIsCurrentlyActive = (tab === cachedTabBtn && activeMode === "cached") ||
                    (tab === drawTabBtn && activeMode === "draw") ||
                    (tab === uploadTabBtn && activeMode === "upload");
                if (!tabIsCurrentlyActive) {
                    tab.style.color = "#0078d4";
                }
            };
            tab.onmouseout = () => {
                const tabIsCurrentlyActive = (tab === cachedTabBtn && activeMode === "cached") ||
                    (tab === drawTabBtn && activeMode === "draw") ||
                    (tab === uploadTabBtn && activeMode === "upload");
                if (!tabIsCurrentlyActive) {
                    tab.style.color = "#605e5c";
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
        const cachedPanel = document.createElement("div");
        cachedPanel.style.display = hasCachedSignature ? "block" : "none";
        cachedPanel.appendChild(cachedNoticeContainer);
        // -----------------------------------------------------------------
        // Draw panel (canvas)
        // -----------------------------------------------------------------
        const drawPanel = document.createElement("div");
        drawPanel.style.display = hasCachedSignature ? "none" : "block";
        const canvasContainer = document.createElement("div");
        canvasContainer.style.backgroundColor = "#ffffff";
        canvasContainer.style.padding = "16px";
        canvasContainer.style.borderRadius = "6px";
        canvasContainer.style.border = "2px dashed #d1d1d1";
        const canvasInstructions = document.createElement("p");
        canvasInstructions.innerText = "Draw your signature below:";
        canvasInstructions.style.fontSize = "12px";
        canvasInstructions.style.color = "#605e5c";
        canvasInstructions.style.margin = "0 0 8px 0";
        canvasInstructions.style.fontWeight = "600";
        const canvas = document.createElement("canvas");
        canvas.width = 560;
        canvas.height = 160;
        canvas.style.border = "1px solid #d1d1d1";
        canvas.style.backgroundColor = "#ffffff";
        canvas.style.cursor = "crosshair";
        canvas.style.display = "block";
        canvas.style.width = "100%";
        canvas.style.borderRadius = "4px";
        canvas.style.boxShadow = "inset 0 1px 3px rgba(0,0,0,0.1)";
        /** Button to clear the drawing canvas. */
        const clearButton = document.createElement("button");
        clearButton.innerText = "Clear Canvas";
        clearButton.style.marginTop = "10px";
        clearButton.style.padding = "8px 16px";
        clearButton.style.fontSize = "12px";
        clearButton.style.backgroundColor = "#f3f2f1";
        clearButton.style.border = "1px solid #d1d1d1";
        clearButton.style.cursor = "pointer";
        clearButton.style.borderRadius = "4px";
        clearButton.style.fontWeight = "600";
        clearButton.style.color = "#323130";
        clearButton.style.transition = "all 0.2s ease";
        clearButton.onmouseover = () => {
            clearButton.style.backgroundColor = "#e1dfdd";
        };
        clearButton.onmouseout = () => {
            clearButton.style.backgroundColor = "#f3f2f1";
        };
        canvasContainer.appendChild(canvasInstructions);
        canvasContainer.appendChild(canvas);
        canvasContainer.appendChild(clearButton);
        drawPanel.appendChild(canvasContainer);
        // -----------------------------------------------------------------
        // Upload panel
        // -----------------------------------------------------------------
        const uploadPanel = document.createElement("div");
        uploadPanel.style.display = "none";
        uploadPanel.style.padding = "32px 24px";
        uploadPanel.style.border = "2px dashed #d1d1d1";
        uploadPanel.style.textAlign = "center";
        uploadPanel.style.backgroundColor = "#ffffff";
        uploadPanel.style.borderRadius = "6px";
        uploadPanel.style.transition = "all 0.2s ease";
        uploadPanel.style.minHeight = "200px";
        uploadPanel.style.flexDirection = "column";
        uploadPanel.style.alignItems = "center";
        uploadPanel.style.justifyContent = "center";
        const uploadIcon = document.createElement("div");
        uploadIcon.innerHTML = "📁";
        uploadIcon.style.fontSize = "48px";
        uploadIcon.style.marginBottom = "16px";
        const uploadText = document.createElement("p");
        uploadText.innerText = "Click to upload or drag and drop";
        uploadText.style.fontSize = "14px";
        uploadText.style.color = "#605e5c";
        uploadText.style.margin = "0 0 8px 0";
        uploadText.style.fontWeight = "600";
        const uploadSubtext = document.createElement("p");
        uploadSubtext.innerText = "PNG, JPG (max 5MB)";
        uploadSubtext.style.fontSize = "12px";
        uploadSubtext.style.color = "#8a8886";
        uploadSubtext.style.margin = "0 0 16px 0";
        const fileInput = document.createElement("input");
        fileInput.type = "file";
        fileInput.accept = "image/png, image/jpeg, image/jpg";
        fileInput.style.fontSize = "12px";
        fileInput.style.marginBottom = "16px";
        const uploadPreview = document.createElement("img");
        uploadPreview.style.maxWidth = "100%";
        uploadPreview.style.height = "auto";
        uploadPreview.style.minHeight = "80px";
        uploadPreview.style.maxHeight = "120px";
        uploadPreview.style.marginTop = "16px";
        uploadPreview.style.display = "none";
        uploadPreview.style.border = "2px solid #0078d4";
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
        let activeMode = hasCachedSignature
            ? "cached"
            : "draw";
        let compressedUploadBase64 = "";
        const ctx = canvas.getContext("2d");
        let isDrawing = false;
        let hasDrawnContent = false;
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
        /**
         * Re-evaluates whether the "I Agree and Sign" button should be
         * enabled, based on the current TFA code entry and selected
         * signature mode.
         */
        const updateButtonState = () => {
            let isCodeValid = true;
            // Check TFA only if it's required
            if (requireTFA) {
                if (verificationInput) {
                    const codeValue = verificationInput.value.trim();
                    isCodeValid =
                        codeValue.length === 5 && codeValue === generatedPasscode;
                    console.log("🔐 TFA Check: Code =", codeValue, "Valid =", isCodeValid);
                }
                else {
                    isCodeValid = false;
                    console.log("🔐 TFA Check: No input element found");
                }
            }
            else {
                console.log("🔐 TFA not required - skipping TFA validation");
            }
            // Check signature validity
            let isSignatureValid = false;
            if (activeMode === "cached") {
                isSignatureValid = hasCachedSignature;
                console.log("✍️ Signature Check (CACHED): Valid =", isSignatureValid);
            }
            else if (activeMode === "draw") {
                const canvasHasContent = ctx !== null && !isCanvasEmpty(ctx, canvas.width, canvas.height);
                isSignatureValid = hasDrawnContent && canvasHasContent;
                console.log("✍️ Signature Check (DRAW): hasDrawnContent =", hasDrawnContent, "canvasHasContent =", canvasHasContent, "Valid =", isSignatureValid);
            }
            else {
                isSignatureValid = compressedUploadBase64.trim() !== "";
                console.log("✍️ Signature Check (UPLOAD): compressedSize =", compressedUploadBase64.length, "Valid =", isSignatureValid);
            }
            console.log("🔘 FINAL STATE: Code Valid =", isCodeValid, "Sig Valid =", isSignatureValid, "Should Enable =", isCodeValid && isSignatureValid);
            if (isCodeValid && isSignatureValid) {
                signButton.disabled = false;
                signButton.style.backgroundColor = "#0078d4";
                signButton.style.cursor = "pointer";
                signButton.style.opacity = "1";
                console.log("✅ Button ENABLED");
            }
            else {
                signButton.disabled = true;
                signButton.style.backgroundColor = "#c8c6c4";
                signButton.style.cursor = "not-allowed";
                signButton.style.opacity = "0.6";
                console.log("❌ Button DISABLED");
            }
        };
        // -----------------------------------------------------------------
        // Tab-switching logic
        // -----------------------------------------------------------------
        /**
         * Activates the given tab, hiding the other panels and updating
         * visual tab styles.
         *
         * @param mode - Which panel to show.
         */
        const setActiveTab = (mode) => {
            activeMode = mode;
            cachedPanel.style.display = mode === "cached" ? "block" : "none";
            drawPanel.style.display = mode === "draw" ? "block" : "none";
            uploadPanel.style.display = mode === "upload" ? "flex" : "none";
            [cachedTabBtn, drawTabBtn, uploadTabBtn].forEach((btn) => {
                btn.style.color = "#605e5c";
                btn.style.borderBottom = "3px solid transparent";
            });
            const activeBtn = mode === "cached"
                ? cachedTabBtn
                : mode === "draw"
                    ? drawTabBtn
                    : uploadTabBtn;
            activeBtn.style.color = "#0078d4";
            activeBtn.style.borderBottom = "3px solid #0078d4";
            console.log("Tab switched to:", mode);
            updateButtonState();
        };
        if (hasCachedSignature) {
            cachedTabBtn.onclick = () => setActiveTab("cached");
        }
        drawTabBtn.onclick = () => setActiveTab("draw");
        uploadTabBtn.onclick = () => setActiveTab("upload");
        /** Removes the cached signature and switches to the draw tab. */
        removeCachedBtn.onclick = () => {
            localStorage.removeItem(STORAGE_KEY);
            hasCachedSignature = false;
            cachedTabBtn.style.display = "none";
            setActiveTab("draw");
        };
        // -----------------------------------------------------------------
        // Two-factor: send-code logic
        // -----------------------------------------------------------------
        if (requireTFA && context.channel && sendCodeButton && verificationInput) {
            let cooldownTimer = null;
            const scBtnRef = sendCodeButton;
            const vInputRef = verificationInput;
            /**
             * Starts a 60-second cooldown on the "Resend Code" button to
             * prevent rapid re-sends.
             */
            const startCooldownTimer = () => {
                let secondsLeft = 60;
                scBtnRef.disabled = true;
                scBtnRef.style.backgroundColor = "#f3f2f1";
                scBtnRef.style.color = "#a19f9d";
                scBtnRef.style.borderColor = "#d1d1d1";
                scBtnRef.style.cursor = "not-allowed";
                scBtnRef.innerText = `Resend (${secondsLeft}s)`;
                if (cooldownTimer) {
                    clearInterval(cooldownTimer);
                }
                cooldownTimer = window.setInterval(() => {
                    secondsLeft -= 1;
                    if (secondsLeft > 0) {
                        scBtnRef.innerText = `Resend (${secondsLeft}s)`;
                    }
                    else {
                        if (cooldownTimer) {
                            clearInterval(cooldownTimer);
                        }
                        scBtnRef.disabled = false;
                        scBtnRef.style.backgroundColor = "#ffffff";
                        scBtnRef.style.color = "#0078d4";
                        scBtnRef.style.borderColor = "#0078d4";
                        scBtnRef.style.cursor = "pointer";
                        scBtnRef.innerText = "Resend Code";
                    }
                }, 1000);
            };
            /**
             * Dispatches the verification passcode through the configured
             * channel (email / Teams) and starts the cooldown timer.
             */
            const triggerSendCode = async () => {
                scBtnRef.innerText = "Sending...";
                scBtnRef.disabled = true;
                try {
                    if (context.spContext && context.channel) {
                        const vResult = await (0, OtpService_1.dispatchOtpPasscode)(context.spContext, {
                            title: context.signer,
                            passcode: generatedPasscode,
                            channel: context.channel,
                        });
                        if (vResult.success && vResult.itemId) {
                            storedVerificationItemId = vResult.itemId;
                            console.log("✅ OTP dispatched successfully");
                        }
                        else {
                            console.error("❌ OTP dispatch failed:", vResult.error);
                            alert("Failed to dispatch verification code. Please try again.");
                        }
                    }
                }
                catch (err) {
                    console.error("Failed to send code:", err);
                    alert("Failed to send verification code. Please try again.");
                }
                finally {
                    startCooldownTimer();
                }
            };
            scBtnRef.onclick = async () => {
                await triggerSendCode();
            };
            vInputRef.oninput = () => {
                console.log("🔐 TFA input changed, checking button state");
                updateButtonState();
            };
            void triggerSendCode();
        }
        else if (requireTFA && !context.channel) {
            console.warn("⚠️ TFA required but no channel provided");
        }
        // -----------------------------------------------------------------
        // Canvas drawing helpers
        // -----------------------------------------------------------------
        /**
         * Translates a mouse or touch event into canvas-relative
         * coordinates, accounting for CSS scaling.
         *
         * @param e - The originating mouse or touch event.
         * @returns `{ x, y }` in canvas-pixel space.
         */
        const getPos = (e) => {
            const rect = canvas.getBoundingClientRect();
            let clientX;
            let clientY;
            if ("touches" in e && e.touches.length > 0) {
                clientX = e.touches[0].clientX;
                clientY = e.touches[0].clientY;
            }
            else if ("changedTouches" in e && e.changedTouches.length > 0) {
                clientX = e.changedTouches[0].clientX;
                clientY = e.changedTouches[0].clientY;
            }
            else {
                clientX = e.clientX;
                clientY = e.clientY;
            }
            return {
                x: (clientX - rect.left) * (canvas.width / rect.width),
                y: (clientY - rect.top) * (canvas.height / rect.height),
            };
        };
        /**
         * Begins a new drawing stroke at the pointer position.
         *
         * @param e - The initiating mouse/touch event.
         */
        const startDraw = (e) => {
            isDrawing = true;
            hasDrawnContent = true;
            const pos = getPos(e);
            ctx === null || ctx === void 0 ? void 0 : ctx.beginPath();
            ctx === null || ctx === void 0 ? void 0 : ctx.moveTo(pos.x, pos.y);
            e.preventDefault();
        };
        /**
         * Extends the current stroke to the pointer's new position.
         *
         * @param e - The move event.
         */
        const drawLine = (e) => {
            if (!isDrawing || !ctx)
                return;
            const pos = getPos(e);
            ctx.lineTo(pos.x, pos.y);
            ctx.stroke();
            console.log("✏️ Drawing, updating button state");
            updateButtonState();
            e.preventDefault();
        };
        /** Ends the current drawing stroke. */
        const stopDraw = () => {
            isDrawing = false;
            console.log("✋ Drawing stopped, updating button state");
            updateButtonState();
        };
        canvas.addEventListener("mousedown", startDraw);
        canvas.addEventListener("mousemove", drawLine);
        window.addEventListener("mouseup", stopDraw);
        canvas.addEventListener("touchstart", startDraw, { passive: false });
        canvas.addEventListener("touchmove", drawLine, { passive: false });
        window.addEventListener("touchend", stopDraw);
        /** Clears the drawing canvas back to a blank white rectangle. */
        clearButton.onclick = () => {
            if (ctx) {
                ctx.fillStyle = "#ffffff";
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                hasDrawnContent = false;
                console.log("🧹 Canvas cleared, updating button state");
                updateButtonState();
            }
        };
        // -----------------------------------------------------------------
        // File upload handler
        // -----------------------------------------------------------------
        /**
         * Handles image file selection: validates size, resizes if needed,
         * and stores the result as a data-URI.
         */
        fileInput.onchange = (e) => {
            const target = e.target;
            if (target.files && target.files[0]) {
                const file = target.files[0];
                /* Enforce the advertised 5 MB limit. */
                if (file.size > MAX_UPLOAD_BYTES) {
                    alert(`File size exceeds 5 MB (${(file.size / 1024 / 1024).toFixed(1)} MB). Please choose a smaller image.`);
                    target.value = "";
                    return;
                }
                const reader = new FileReader();
                reader.onload = (uploadEvent) => {
                    var _a;
                    if ((_a = uploadEvent.target) === null || _a === void 0 ? void 0 : _a.result) {
                        const img = new Image();
                        img.onload = () => {
                            const tempCanvas = document.createElement("canvas");
                            const maxW = 560;
                            const maxH = 160;
                            let w = img.width;
                            let h = img.height;
                            if (w > maxW) {
                                h = Math.round((h * maxW) / w);
                                w = maxW;
                            }
                            if (h > maxH) {
                                w = Math.round((w * maxH) / h);
                                h = maxH;
                            }
                            tempCanvas.width = w;
                            tempCanvas.height = h;
                            const tCtx = tempCanvas.getContext("2d");
                            if (tCtx) {
                                tCtx.fillStyle = "#ffffff";
                                tCtx.fillRect(0, 0, w, h);
                                tCtx.drawImage(img, 0, 0, w, h);
                                compressedUploadBase64 =
                                    tempCanvas.toDataURL("image/png");
                                uploadPreview.src = compressedUploadBase64;
                                uploadPreview.style.display = "block";
                                console.log("📤 File uploaded, updating button state");
                                updateButtonState();
                            }
                        };
                        img.src = uploadEvent.target.result;
                    }
                };
                reader.readAsDataURL(file);
            }
        };
        // -----------------------------------------------------------------
        // Footer (Cancel + Sign buttons)
        // -----------------------------------------------------------------
        const footerSection = document.createElement("div");
        footerSection.style.padding = "20px 24px";
        footerSection.style.backgroundColor = "#ffffff";
        footerSection.style.borderTop = "1px solid #edebe9";
        footerSection.style.display = "flex";
        footerSection.style.justifyContent = "flex-end";
        footerSection.style.gap = "12px";
        const cancelButton = document.createElement("button");
        cancelButton.innerText = "Cancel";
        cancelButton.style.padding = "10px 24px";
        cancelButton.style.backgroundColor = "#ffffff";
        cancelButton.style.border = "1px solid #8a8886";
        cancelButton.style.cursor = "pointer";
        cancelButton.style.fontSize = "14px";
        cancelButton.style.fontWeight = "600";
        cancelButton.style.color = "#323130";
        cancelButton.style.borderRadius = "4px";
        cancelButton.style.transition = "all 0.2s ease";
        cancelButton.onmouseover = () => {
            cancelButton.style.backgroundColor = "#f3f2f1";
        };
        cancelButton.onmouseout = () => {
            cancelButton.style.backgroundColor = "#ffffff";
        };
        signButton.onmouseover = () => {
            if (!signButton.disabled) {
                signButton.style.backgroundColor = "#005a9e";
            }
        };
        signButton.onmouseout = () => {
            if (!signButton.disabled) {
                signButton.style.backgroundColor = "#0078d4";
            }
        };
        // -----------------------------------------------------------------
        // Cleanup helper
        // -----------------------------------------------------------------
        /**
         * Removes the modal overlay from the DOM and unregisters window-
         * level event listeners to prevent memory leaks.
         */
        const cleanup = () => {
            window.removeEventListener("mouseup", stopDraw);
            window.removeEventListener("touchend", stopDraw);
            document.body.removeChild(overlay);
        };
        /** Cancel button dismisses the dialog and resolves with `undefined`. */
        cancelButton.onclick = () => {
            cleanup();
            resolve(undefined);
        };
        // -----------------------------------------------------------------
        // Sign button handler
        // -----------------------------------------------------------------
        /**
         * Validates the TFA code (if required), extracts the signature
         * data from the active panel, compresses it, optionally caches it,
         * then generates the SHA-256 audit record.
         */
        signButton.onclick = async () => {
            if (requireTFA && verificationInput) {
                const enteredCode = verificationInput.value.trim();
                if (enteredCode !== generatedPasscode) {
                    alert("Invalid verification code. Please enter the correct 5-digit code.");
                    return;
                }
            }
            let finalDataUrl = "";
            if (activeMode === "cached") {
                if (!cachedCompressedSig) {
                    alert("No cached signature found. Please select another option.");
                    return;
                }
                finalDataUrl = lzwDecompress(cachedCompressedSig);
            }
            else if (activeMode === "draw") {
                if (!ctx ||
                    isCanvasEmpty(ctx, canvas.width, canvas.height)) {
                    alert("Please draw your signature before proceeding.");
                    return;
                }
                finalDataUrl = downscaleCanvas(canvas, 560, 160);
            }
            else {
                if (!compressedUploadBase64 ||
                    compressedUploadBase64.trim() === "") {
                    alert("Please upload a signature image before proceeding.");
                    return;
                }
                finalDataUrl = compressedUploadBase64;
            }
            const compressedString = lzwCompress(finalDataUrl);
            if (cacheCheckbox.checked) {
                localStorage.setItem(STORAGE_KEY, compressedString);
            }
            else {
                localStorage.removeItem(STORAGE_KEY);
            }
            cleanup();
            try {
                const fullContext = Object.assign(Object.assign({}, context), { signatureData: compressedString });
                const result = await generateSecureAuditRecordInternal(fullContext);
                if (storedVerificationItemId !== undefined) {
                    result.verificationItemId = storedVerificationItemId;
                }
                resolve(result);
            }
            catch (error) {
                console.error("Signing failed:", error);
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
exports.promptAndGenerateSecureAudit = promptAndGenerateSecureAudit;
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
function getReportableSignature(compressedSignatureData) {
    const decompressed = lzwDecompress(compressedSignatureData);
    if (!decompressed || !decompressed.startsWith("data:image")) {
        return "";
    }
    return decompressed;
}
exports.getReportableSignature = getReportableSignature;
// ---------------------------------------------------------------------------
// LZW compression / decompression
// ---------------------------------------------------------------------------
/**
 * Compresses a string using the LZW (Lempel-Ziv-Welch) algorithm.
 *
 * **Caveat:** When the dictionary grows past 65 535 entries,
 * `String.fromCharCode` will produce values that may not survive
 * a round-trip through `localStorage` or JSON.  For very large
 * inputs consider chunking or switching to a Uint16Array-backed
 * encoding.
 *
 * @param input - The raw string to compress.
 * @returns The compressed string (each character encodes one
 *          dictionary index).
 */
function lzwCompress(input) {
    const dictionary = {};
    let c = "";
    let wc = "";
    let w = "";
    const result = [];
    let dictionarySize = 256;
    for (let i = 0; i < 256; i += 1) {
        dictionary[String.fromCharCode(i)] = i;
    }
    for (let i = 0; i < input.length; i += 1) {
        c = input.charAt(i);
        wc = w + c;
        if (Object.prototype.hasOwnProperty.call(dictionary, wc)) {
            w = wc;
        }
        else {
            result.push(dictionary[w]);
            dictionary[wc] = dictionarySize++;
            w = String(c);
        }
    }
    if (w !== "") {
        result.push(dictionary[w]);
    }
    return result
        .map((n) => String.fromCharCode(n))
        .join("");
}
/**
 * Decompresses a string that was compressed with {@link lzwCompress}.
 *
 * @param compressed - The LZW-compressed string.
 * @returns The original string, or an empty string if the input is
 *          empty or contains an unrecognised dictionary reference.
 */
function lzwDecompress(compressed) {
    const dictionary = {};
    const result = [];
    let dictionarySize = 256;
    for (let i = 0; i < 256; i += 1) {
        dictionary[i] = String.fromCharCode(i);
    }
    if (compressed.length === 0)
        return "";
    let w = String.fromCharCode(compressed.charCodeAt(0));
    result.push(w);
    let entry = "";
    for (let i = 1; i < compressed.length; i += 1) {
        const k = compressed.charCodeAt(i);
        if (dictionary[k] !== undefined) {
            entry = dictionary[k];
        }
        else if (k === dictionarySize) {
            entry = w + w.charAt(0);
        }
        else {
            console.warn(`lzwDecompress: unrecognised dictionary index ${k} at position ${i}.`);
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
function downscaleCanvas(sourceCanvas, targetWidth, targetHeight) {
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
function isCanvasEmpty(ctx, width, height) {
    const pixelBuffer = ctx.getImageData(0, 0, width, height).data;
    for (let i = 0; i < pixelBuffer.length; i += 4) {
        if (pixelBuffer[i] !== 255 ||
            pixelBuffer[i + 1] !== 255 ||
            pixelBuffer[i + 2] !== 255) {
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
 * **Note:** Only top-level payload keys are sorted. If your payload
 * contains nested objects whose key order may vary, consider using a
 * deep-sort utility before calling this function.
 *
 * @param auditRecord - The SharePointAuditRecord returned from `promptAndGenerateSecureAudit`.
 * @param signer      - The signer's email address or display name (must match original signer).
 * @param payload     - The original payload object that was signed.
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
async function verifySecureAuditRecord(auditRecord, signer, payload) {
    try {
        if (!auditRecord || !signer || !payload) {
            console.warn("verifySecureAuditRecord: Missing required parameters", {
                auditRecord: !!auditRecord,
                signer: !!signer,
                payload: !!payload,
            });
            return false;
        }
        // Sort payload keys for consistent hashing
        const sortedPayloadString = JSON.stringify(payload, Object.keys(payload).sort());
        // Reconstruct the canonical audit envelope
        const auditEnvelope = {
            payload: JSON.parse(sortedPayloadString),
            signer: signer.toLowerCase().trim(),
            timestamp: auditRecord.signatureTimestamp.trim(),
        };
        // Hash the audit envelope
        const auditEnvelopeJson = JSON.stringify(auditEnvelope);
        const encoder = new TextEncoder();
        const encodedBytes = encoder.encode(auditEnvelopeJson);
        const hashBuffer = await window.crypto.subtle.digest("SHA-256", encodedBytes.buffer);
        const hashArray = Array.from(new Uint8Array(hashBuffer));
        const recomputedHash = hashArray
            .map((b) => b.toString(16).padStart(2, "0"))
            .join("");
        // Compare hashes
        const isValid = recomputedHash === auditRecord.signatureHash.trim();
        console.log("verifySecureAuditRecord: Verification result =", isValid, "Recomputed hash =", recomputedHash.substring(0, 16) + "...");
        return isValid;
    }
    catch (error) {
        console.error("verifySecureAuditRecord: Error during verification", error);
        return false;
    }
}
exports.verifySecureAuditRecord = verifySecureAuditRecord;
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
 * @param context - The full signer context, extended with the
 *                  compressed signature data string.
 * @returns A fully populated `SharePointAuditRecord`.
 *
 * @throws {Error} If the payload is empty or the signer is blank.
 */
async function generateSecureAuditRecordInternal(context) {
    const { payload, signer, signatureData } = context;
    if (!payload || Object.keys(payload).length === 0) {
        throw new Error("Payload cannot be empty.");
    }
    if (!signer || signer.trim() === "") {
        throw new Error("Signer identifier is required for sealing.");
    }
    const signatureTimestamp = new Date().toISOString();
    const sortedPayloadString = JSON.stringify(payload, Object.keys(payload).sort());
    const auditEnvelope = {
        payload: JSON.parse(sortedPayloadString),
        signer: signer.toLowerCase().trim(),
        timestamp: signatureTimestamp.trim(),
    };
    const auditEnvelopeJson = JSON.stringify(auditEnvelope);
    const encoder = new TextEncoder();
    const encodedBytes = encoder.encode(auditEnvelopeJson);
    // Passing `encodedBytes.buffer as ArrayBuffer` satisfies the DOM BufferSource definition
    const hashBuffer = await window.crypto.subtle.digest("SHA-256", encodedBytes.buffer);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    const signatureHash = hashArray
        .map((b) => b.toString(16).padStart(2, "0"))
        .join("");
    return {
        signatureHash,
        signatureData,
        signatureTimestamp,
        verificationItemId: context.itemID,
    };
}
