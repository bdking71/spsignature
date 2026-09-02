import { WebPartContext } from "@microsoft/sp-webpart-base";
import { createPendingVerification, DeliveryChannel } from "./VerificationService";

const STORAGE_KEY = "secure_audit_cached_signature_v1";

export interface SignerContext {
  itemID: number;
  signer: string;
  payload: Record<string, unknown>;
  spContext: WebPartContext;
  channel?: DeliveryChannel;
  requireTFA?: boolean;
}

export interface SharePointAuditRecord {
  signatureHash: string;
  signatureData: string;
  signatureTimestamp: string;
  verificationItemId: number;
}

export interface AuditEnvelopeRecord {
  payload: Record<string, unknown>;
  signer: string;
  timestamp: string;
}

function generateFiveDigitPasscode(): string {
  const array = new Uint32Array(1);
  window.crypto.getRandomValues(array);
  const code = (array[0] % 90000) + 10000;
  return code.toString();
}

export async function promptAndGenerateSecureAudit(
  context: SignerContext,
  modalTitle: string = "Approve Purchase Requisition",
  warningMessage: string = "Please review your entry and apply your signature to finalize this record."
): Promise<SharePointAuditRecord | undefined> {
  if (!context.spContext) {
    throw new Error("Execution restricted: Valid WebPartContext must be supplied.");
  }

  return new Promise<SharePointAuditRecord | undefined>((resolve): void => {
    const requireTFA = context.requireTFA !== false;
    const generatedPasscode = generateFiveDigitPasscode();
    let storedVerificationItemId: number = context.itemID;

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
    overlay.style.fontFamily = "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";
    overlay.style.backdropFilter = "blur(3px)";

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

    const headerSection: HTMLDivElement = document.createElement("div");
    headerSection.style.backgroundColor = "#0078d4";
    headerSection.style.padding = "20px 24px";
    headerSection.style.borderBottom = "3px solid #005a9e";

    const title: HTMLHeadingElement = document.createElement("h2");
    title.innerText = modalTitle;
    title.style.margin = "0";
    title.style.color = "#ffffff";
    title.style.fontSize = "20px";
    title.style.fontWeight = "600";
    title.style.letterSpacing = "0.3px";

    headerSection.appendChild(title);

    const contentSection: HTMLDivElement = document.createElement("div");
    contentSection.style.padding = "24px";
    contentSection.style.backgroundColor = "#fafafa";
    contentSection.style.overflowY = "auto";
    contentSection.style.flex = "1";

    const body: HTMLParagraphElement = document.createElement("p");
    body.innerHTML = warningMessage;
    body.style.fontSize = "14px";
    body.style.color = "#323130";
    body.style.lineHeight = "1.6";
    body.style.margin = "0 0 20px 0";

    const signButton: HTMLButtonElement = document.createElement("button");
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

    const verificationContainer: HTMLDivElement = document.createElement("div");
    if (requireTFA) {
      verificationContainer.style.marginBottom = "20px";
      verificationContainer.style.padding = "16px";
      verificationContainer.style.backgroundColor = "#fff8e1";
      verificationContainer.style.border = "1px solid #ffd54f";
      verificationContainer.style.borderRadius = "6px";
      verificationContainer.style.boxShadow = "0 2px 4px rgba(0,0,0,0.05)";

      const verificationHeader: HTMLDivElement = document.createElement("div");
      verificationHeader.style.display = "flex";
      verificationHeader.style.alignItems = "center";
      verificationHeader.style.marginBottom = "12px";
      verificationHeader.style.gap = "8px";

      const lockIcon: HTMLSpanElement = document.createElement("span");
      lockIcon.innerHTML = "🔐";
      lockIcon.style.fontSize = "18px";

      const verificationLabel: HTMLLabelElement = document.createElement("label");
      verificationLabel.innerText = "Two-Factor Authentication";
      verificationLabel.style.display = "block";
      verificationLabel.style.fontSize = "13px";
      verificationLabel.style.fontWeight = "700";
      verificationLabel.style.color = "#323130";
      verificationLabel.style.margin = "0";

      verificationHeader.appendChild(lockIcon);
      verificationHeader.appendChild(verificationLabel);

      const verificationDescription: HTMLParagraphElement = document.createElement("p");
      verificationDescription.innerText = `A 5-digit verification code has been sent to ${context.channel === "teams" ? "Microsoft Teams" : "your email"}.`;
      verificationDescription.style.fontSize = "12px";
      verificationDescription.style.color = "#605e5c";
      verificationDescription.style.margin = "0 0 12px 0";
      verificationDescription.style.lineHeight = "1.5";

      const verificationRow: HTMLDivElement = document.createElement("div");
      verificationRow.style.display = "flex";
      verificationRow.style.gap = "10px";

      const verificationInput: HTMLInputElement = document.createElement("input");
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

      verificationInput.onfocus = (): void => {
        verificationInput.style.borderColor = "#0078d4";
      };
      verificationInput.onblur = (): void => {
        verificationInput.style.borderColor = "#d1d1d1";
      };

      const sendCodeButton: HTMLButtonElement = document.createElement("button");
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

      sendCodeButton.onmouseover = (): void => {
        sendCodeButton.style.backgroundColor = "#0078d4";
        sendCodeButton.style.color = "#ffffff";
      };
      sendCodeButton.onmouseout = (): void => {
        if (!sendCodeButton.disabled) {
          sendCodeButton.style.backgroundColor = "#ffffff";
          sendCodeButton.style.color = "#0078d4";
        }
      };

      verificationRow.appendChild(verificationInput);
      verificationRow.appendChild(sendCodeButton);

      verificationContainer.appendChild(verificationHeader);
      verificationContainer.appendChild(verificationDescription);
      verificationContainer.appendChild(verificationRow);
    }

    const cachedCompressedSig: string | null = localStorage.getItem(STORAGE_KEY);
    const hasCachedSignature: boolean = cachedCompressedSig !== null && cachedCompressedSig.trim() !== "";
    const decompressedCachedDataUri: string = (hasCachedSignature && cachedCompressedSig)
      ? lzwDecompress(cachedCompressedSig)
      : "";

    const signatureSection: HTMLDivElement = document.createElement("div");
    signatureSection.style.marginBottom = "20px";
    signatureSection.style.padding = "16px";
    signatureSection.style.backgroundColor = "#ffffff";
    signatureSection.style.border = "1px solid #d1d1d1";
    signatureSection.style.borderRadius = "6px";
    signatureSection.style.boxShadow = "0 2px 4px rgba(0,0,0,0.05)";

    const signatureHeader: HTMLDivElement = document.createElement("div");
    signatureHeader.style.marginBottom = "12px";

    const signatureLabel: HTMLHeadingElement = document.createElement("h3");
    signatureLabel.innerText = "Digital Signature";
    signatureLabel.style.fontSize = "14px";
    signatureLabel.style.fontWeight = "700";
    signatureLabel.style.color = "#323130";
    signatureLabel.style.margin = "0 0 4px 0";

    const signatureSubtext: HTMLParagraphElement = document.createElement("p");
    signatureSubtext.innerText = "Choose how you'd like to provide your signature";
    signatureSubtext.style.fontSize = "12px";
    signatureSubtext.style.color = "#605e5c";
    signatureSubtext.style.margin = "0";

    signatureHeader.appendChild(signatureLabel);
    signatureHeader.appendChild(signatureSubtext);

    const cacheControlBox: HTMLDivElement = document.createElement("div");
    cacheControlBox.style.marginTop = "16px";
    cacheControlBox.style.padding = "12px";
    cacheControlBox.style.backgroundColor = "#f3f2f1";
    cacheControlBox.style.border = "1px solid #d1d1d1";
    cacheControlBox.style.borderRadius = "4px";
    cacheControlBox.style.fontSize = "12px";

    const cacheCheckboxLabel: HTMLLabelElement = document.createElement("label");
    cacheCheckboxLabel.style.display = "flex";
    cacheCheckboxLabel.style.alignItems = "center";
    cacheCheckboxLabel.style.gap = "8px";
    cacheCheckboxLabel.style.cursor = "pointer";

    const cacheCheckbox: HTMLInputElement = document.createElement("input");
    cacheCheckbox.type = "checkbox";
    cacheCheckbox.checked = true;
    cacheCheckbox.style.cursor = "pointer";
    cacheCheckbox.style.width = "16px";
    cacheCheckbox.style.height = "16px";

    const cacheCheckboxText: HTMLSpanElement = document.createElement("span");
    cacheCheckboxText.innerText = "Remember my signature on this device for future transactions";
    cacheCheckboxText.style.color = "#323130";
    cacheCheckboxText.style.fontSize = "12px";

    cacheCheckboxLabel.appendChild(cacheCheckbox);
    cacheCheckboxLabel.appendChild(cacheCheckboxText);
    cacheControlBox.appendChild(cacheCheckboxLabel);

    const cachedNoticeContainer: HTMLDivElement = document.createElement("div");
    cachedNoticeContainer.style.marginBottom = "16px";
    cachedNoticeContainer.style.padding = "16px";
    cachedNoticeContainer.style.backgroundColor = "#e6f4ff";
    cachedNoticeContainer.style.border = "1px solid #0078d4";
    cachedNoticeContainer.style.borderRadius = "6px";
    cachedNoticeContainer.style.display = hasCachedSignature ? "block" : "none";

    const cachedNoticeText: HTMLDivElement = document.createElement("div");
    cachedNoticeText.innerText = "✓ Saved signature on file";
    cachedNoticeText.style.fontSize = "13px";
    cachedNoticeText.style.marginBottom = "12px";
    cachedNoticeText.style.fontWeight = "600";
    cachedNoticeText.style.color = "#0078d4";

    const cachedPreviewImage: HTMLImageElement = document.createElement("img");
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

    const removeCachedBtn: HTMLButtonElement = document.createElement("button");
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

    removeCachedBtn.onmouseover = (): void => {
      removeCachedBtn.style.backgroundColor = "#a4262c";
      removeCachedBtn.style.color = "#ffffff";
    };
    removeCachedBtn.onmouseout = (): void => {
      removeCachedBtn.style.backgroundColor = "#ffffff";
      removeCachedBtn.style.color = "#a4262c";
    };

    cachedNoticeContainer.appendChild(cachedNoticeText);
    cachedNoticeContainer.appendChild(cachedPreviewImage);
    cachedNoticeContainer.appendChild(removeCachedBtn);

    const modeContainer: HTMLDivElement = document.createElement("div");
    modeContainer.style.display = "flex";
    modeContainer.style.gap = "8px";
    modeContainer.style.marginBottom = "16px";
    modeContainer.style.borderBottom = "2px solid #edebe9";

    const createTab = (text: string, isActive: boolean): HTMLButtonElement => {
      const tab = document.createElement("button");
      tab.innerText = text;
      tab.style.padding = "10px 16px";
      tab.style.fontSize = "13px";
      tab.style.cursor = "pointer";
      tab.style.fontWeight = "600";
      tab.style.backgroundColor = "transparent";
      tab.style.color = isActive ? "#0078d4" : "#605e5c";
      tab.style.border = "none";
      tab.style.borderBottom = isActive ? "3px solid #0078d4" : "3px solid transparent";
      tab.style.transition = "all 0.2s ease";
      tab.style.outline = "none";

      tab.onmouseover = (): void => {
        if (!isActive) {
          tab.style.color = "#0078d4";
        }
      };
      tab.onmouseout = (): void => {
        if (!isActive) {
          tab.style.color = "#605e5c";
        }
      };

      return tab;
    };

    const cachedTabBtn = createTab("Saved", hasCachedSignature);
    cachedTabBtn.style.display = hasCachedSignature ? "inline-block" : "none";

    const drawTabBtn = createTab("Draw New", !hasCachedSignature);
    const uploadTabBtn = createTab("Upload Image", false);

    if (hasCachedSignature) {
      modeContainer.appendChild(cachedTabBtn);
    }
    modeContainer.appendChild(drawTabBtn);
    modeContainer.appendChild(uploadTabBtn);

    const cachedPanel: HTMLDivElement = document.createElement("div");
    cachedPanel.style.display = hasCachedSignature ? "block" : "none";
    cachedPanel.appendChild(cachedNoticeContainer);

    const drawPanel: HTMLDivElement = document.createElement("div");
    drawPanel.style.display = hasCachedSignature ? "none" : "block";

    const canvasContainer: HTMLDivElement = document.createElement("div");
    canvasContainer.style.backgroundColor = "#ffffff";
    canvasContainer.style.padding = "16px";
    canvasContainer.style.borderRadius = "6px";
    canvasContainer.style.border = "2px dashed #d1d1d1";

    const canvasInstructions: HTMLParagraphElement = document.createElement("p");
    canvasInstructions.innerText = "Draw your signature below:";
    canvasInstructions.style.fontSize = "12px";
    canvasInstructions.style.color = "#605e5c";
    canvasInstructions.style.margin = "0 0 8px 0";
    canvasInstructions.style.fontWeight = "600";

    const canvas: HTMLCanvasElement = document.createElement("canvas");
    canvas.width = 560;
    canvas.height = 160;
    canvas.style.border = "1px solid #d1d1d1";
    canvas.style.backgroundColor = "#ffffff";
    canvas.style.cursor = "crosshair";
    canvas.style.display = "block";
    canvas.style.width = "100%";
    canvas.style.borderRadius = "4px";
    canvas.style.boxShadow = "inset 0 1px 3px rgba(0,0,0,0.1)";

    const clearButton: HTMLButtonElement = document.createElement("button");
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

    clearButton.onmouseover = (): void => {
      clearButton.style.backgroundColor = "#e1dfdd";
    };
    clearButton.onmouseout = (): void => {
      clearButton.style.backgroundColor = "#f3f2f1";
    };

    canvasContainer.appendChild(canvasInstructions);
    canvasContainer.appendChild(canvas);
    canvasContainer.appendChild(clearButton);
    drawPanel.appendChild(canvasContainer);

    const uploadPanel: HTMLDivElement = document.createElement("div");
    uploadPanel.style.display = "none";
    uploadPanel.style.padding = "32px 24px";
    uploadPanel.style.border = "2px dashed #d1d1d1";
    uploadPanel.style.textAlign = "center";
    uploadPanel.style.backgroundColor = "#ffffff";
    uploadPanel.style.borderRadius = "6px";
    uploadPanel.style.transition = "all 0.2s ease";
    uploadPanel.style.minHeight = "200px";
    uploadPanel.style.display = "none";
    uploadPanel.style.flexDirection = "column";
    uploadPanel.style.alignItems = "center";
    uploadPanel.style.justifyContent = "center";

    const uploadIcon: HTMLDivElement = document.createElement("div");
    uploadIcon.innerHTML = "📁";
    uploadIcon.style.fontSize = "48px";
    uploadIcon.style.marginBottom = "16px";

    const uploadText: HTMLParagraphElement = document.createElement("p");
    uploadText.innerText = "Click to upload or drag and drop";
    uploadText.style.fontSize = "14px";
    uploadText.style.color = "#605e5c";
    uploadText.style.margin = "0 0 8px 0";
    uploadText.style.fontWeight = "600";

    const uploadSubtext: HTMLParagraphElement = document.createElement("p");
    uploadSubtext.innerText = "PNG, JPG (max 5MB)";
    uploadSubtext.style.fontSize = "12px";
    uploadSubtext.style.color = "#8a8886";
    uploadSubtext.style.margin = "0 0 16px 0";

    const fileInput: HTMLInputElement = document.createElement("input");
    fileInput.type = "file";
    fileInput.accept = "image/png, image/jpeg, image/jpg";
    fileInput.style.fontSize = "12px";
    fileInput.style.marginBottom = "16px";

    const uploadPreview: HTMLImageElement = document.createElement("img");
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

    let activeMode: "cached" | "draw" | "upload" = hasCachedSignature ? "cached" : "draw";
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

    const updateButtonState = (): void => {
      let isCodeValid = true;

      if (requireTFA) {
        const verificationInput = verificationContainer.querySelector("input") as HTMLInputElement;
        const codeValue = verificationInput.value.trim();
        isCodeValid = codeValue.length === 5 && codeValue === generatedPasscode;
      }

      let isSignatureValid = false;
      if (activeMode === "cached") {
        isSignatureValid = hasCachedSignature;
      } else if (activeMode === "draw") {
        isSignatureValid = hasDrawnContent && ctx !== null && !isCanvasEmpty(ctx, canvas.width, canvas.height);
      } else {
        isSignatureValid = compressedUploadBase64.trim() !== "";
      }

      if (isCodeValid && isSignatureValid) {
        signButton.disabled = false;
        signButton.style.backgroundColor = "#0078d4";
        signButton.style.cursor = "pointer";
        signButton.style.opacity = "1";
      } else {
        signButton.disabled = true;
        signButton.style.backgroundColor = "#c8c6c4";
        signButton.style.cursor = "not-allowed";
        signButton.style.opacity = "0.6";
      }
    };

    const setActiveTab = (mode: "cached" | "draw" | "upload"): void => {
      activeMode = mode;
      cachedPanel.style.display = mode === "cached" ? "block" : "none";
      drawPanel.style.display = mode === "draw" ? "block" : "none";
      uploadPanel.style.display = mode === "upload" ? "flex" : "none";

      [cachedTabBtn, drawTabBtn, uploadTabBtn].forEach((btn) => {
        btn.style.color = "#605e5c";
        btn.style.borderBottom = "3px solid transparent";
      });

      const activeBtn = mode === "cached" ? cachedTabBtn : mode === "draw" ? drawTabBtn : uploadTabBtn;
      activeBtn.style.color = "#0078d4";
      activeBtn.style.borderBottom = "3px solid #0078d4";

      updateButtonState();
    };

    if (hasCachedSignature) {
      cachedTabBtn.onclick = (): void => setActiveTab("cached");
    }
    drawTabBtn.onclick = (): void => setActiveTab("draw");
    uploadTabBtn.onclick = (): void => setActiveTab("upload");

    removeCachedBtn.onclick = (): void => {
      localStorage.removeItem(STORAGE_KEY);
      cachedTabBtn.style.display = "none";
      setActiveTab("draw");
    };

    if (requireTFA && context.channel) {
      const sendCodeButton = verificationContainer.querySelector("button") as HTMLButtonElement;
      const verificationInput = verificationContainer.querySelector("input") as HTMLInputElement;

      let cooldownTimer: number | null = null;
      const startCooldownTimer = (): void => {
        let secondsLeft = 60;
        sendCodeButton.disabled = true;
        sendCodeButton.style.backgroundColor = "#f3f2f1";
        sendCodeButton.style.color = "#a19f9d";
        sendCodeButton.style.borderColor = "#d1d1d1";
        sendCodeButton.style.cursor = "not-allowed";
        sendCodeButton.innerText = `Resend (${secondsLeft}s)`;

        if (cooldownTimer) {
          clearInterval(cooldownTimer);
        }

        cooldownTimer = window.setInterval(() => {
          secondsLeft -= 1;
          if (secondsLeft > 0) {
            sendCodeButton.innerText = `Resend (${secondsLeft}s)`;
          } else {
            if (cooldownTimer) {
              clearInterval(cooldownTimer);
            }
            sendCodeButton.disabled = false;
            sendCodeButton.style.backgroundColor = "#ffffff";
            sendCodeButton.style.color = "#0078d4";
            sendCodeButton.style.borderColor = "#0078d4";
            sendCodeButton.style.cursor = "pointer";
            sendCodeButton.innerText = "Resend Code";
          }
        }, 1000);
      };

      const triggerSendCode = async (): Promise<void> => {
        sendCodeButton.innerText = "Sending...";
        sendCodeButton.disabled = true;

        try {
          if (context.spContext && context.channel) {
            const vResult = await createPendingVerification(
              context.spContext,
              {
                title: context.signer,
                passcode: generatedPasscode,
                channel: context.channel,
              }
            );

            if (vResult.success && vResult.itemId) {
              storedVerificationItemId = vResult.itemId;
            } else {
              alert("Failed to dispatch verification code. Please try again.");
            }
          }
        } catch (err) {
          console.error("Failed to send code:", err);
          alert("Failed to send verification code. Please try again.");
        } finally {
          startCooldownTimer();
        }
      };

      sendCodeButton.onclick = async (): Promise<void> => {
        await triggerSendCode();
      };

      verificationInput.oninput = (): void => {
        updateButtonState();
      };

      void triggerSendCode();
    }

    const getPos = (e: MouseEvent | TouchEvent): { x: number; y: number } => {
      const rect: DOMRect = canvas.getBoundingClientRect();
      const clientX = "touches" in e ? e.touches[0].clientX : (e as MouseEvent).clientX;
      const clientY = "touches" in e ? e.touches[0].clientY : (e as MouseEvent).clientY;
      return {
        x: (clientX - rect.left) * (canvas.width / rect.width),
        y: (clientY - rect.top) * (canvas.height / rect.height)
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
    window.addEventListener("mouseup", stopDraw);
    canvas.addEventListener("touchstart", startDraw);
    canvas.addEventListener("touchmove", drawLine);
    window.addEventListener("touchend", stopDraw);

    clearButton.onclick = (): void => {
      if (ctx) {
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        hasDrawnContent = false;
        updateButtonState();
      }
    };

    fileInput.onchange = (e: Event): void => {
      const target = e.target as HTMLInputElement;
      if (target.files && target.files[0]) {
        const file = target.files[0];
        const reader = new FileReader();
        reader.onload = (uploadEvent): void => {
          if (uploadEvent.target?.result) {
            const img = new Image();
            img.onload = (): void => {
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
                compressedUploadBase64 = tempCanvas.toDataURL("image/png");
                uploadPreview.src = compressedUploadBase64;
                uploadPreview.style.display = "block";
                updateButtonState();
              }
            };
            img.src = uploadEvent.target.result as string;
          }
        };
        reader.readAsDataURL(file);
      }
    };

    const footerSection: HTMLDivElement = document.createElement("div");
    footerSection.style.padding = "20px 24px";
    footerSection.style.backgroundColor = "#ffffff";
    footerSection.style.borderTop = "1px solid #edebe9";
    footerSection.style.display = "flex";
    footerSection.style.justifyContent = "flex-end";
    footerSection.style.gap = "12px";

    const cancelButton: HTMLButtonElement = document.createElement("button");
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

    cancelButton.onmouseover = (): void => {
      cancelButton.style.backgroundColor = "#f3f2f1";
    };
    cancelButton.onmouseout = (): void => {
      cancelButton.style.backgroundColor = "#ffffff";
    };

    signButton.onmouseover = (): void => {
      if (!signButton.disabled) {
        signButton.style.backgroundColor = "#005a9e";
      }
    };
    signButton.onmouseout = (): void => {
      if (!signButton.disabled) {
        signButton.style.backgroundColor = "#0078d4";
      }
    };

    const cleanup = (): void => {
      document.body.removeChild(overlay);
    };

    cancelButton.onclick = (): void => {
      cleanup();
      resolve(undefined);
    };

    signButton.onclick = async (): Promise<void> => {
      if (requireTFA) {
        const verificationInput = verificationContainer.querySelector("input") as HTMLInputElement;
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
      } else if (activeMode === "draw") {
        if (!ctx || isCanvasEmpty(ctx, canvas.width, canvas.height)) {
          alert("Please draw your signature before proceeding.");
          return;
        }
        finalDataUrl = downscaleCanvas(canvas, 560, 160);
      } else {
        if (!compressedUploadBase64 || compressedUploadBase64.trim() === "") {
          alert("Please upload a signature image before proceeding.");
          return;
        }
        finalDataUrl = compressedUploadBase64;
      }

      const compressedString = lzwCompress(finalDataUrl);

      if (cacheCheckbox.checked) {
        localStorage.setItem(STORAGE_KEY, compressedString);
      } else {
        localStorage.removeItem(STORAGE_KEY);
      }

      cleanup();

      try {
        const fullContext = {
          ...context,
          signatureData: compressedString,
        };

        const result: SharePointAuditRecord =
          await generateSecureAuditRecordInternal(fullContext);

        if (storedVerificationItemId) {
          result.verificationItemId = storedVerificationItemId;
        }

        resolve(result);
      } catch (error: unknown) {
        console.error("Signing failed:", error);
        resolve(undefined);
      }
    };

    updateButtonState();

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

export function getReportableSignature(compressedSignatureData: string): string {
  const decompressed = lzwDecompress(compressedSignatureData);
  if (!decompressed || !decompressed.startsWith("data:image")) {
    return "";
  }
  return decompressed;
}

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
      dictionary[wc] = dictionarySize++;
      w = String(c);
    }
  }

  if (w !== "") {
    result.push(dictionary[w]);
  }

  return result.map((n: number): string => String.fromCharCode(n)).join("");
}

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
    if (dictionary[k]) {
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

function downscaleCanvas(sourceCanvas: HTMLCanvasElement, targetWidth: number, targetHeight: number): string {
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

function isCanvasEmpty(ctx: CanvasRenderingContext2D, width: number, height: number): boolean {
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

export async function verifySecureAuditRecord(
  payloadToVerify: Record<string, unknown>,
  signer: string,
  timestamp: string,
  compressedSignatureData: string,
  storedHash: string
): Promise<boolean> {
  try {
    if (!payloadToVerify || !storedHash) return false;

    const sortedPayloadString: string = JSON.stringify(
      payloadToVerify,
      Object.keys(payloadToVerify).sort()
    );

    const auditEnvelope: AuditEnvelopeRecord = {
      payload: JSON.parse(sortedPayloadString) as Record<string, unknown>,
      signer: signer.toLowerCase().trim(),
      timestamp: timestamp.trim(),
    };

    const auditEnvelopeJson: string = JSON.stringify(auditEnvelope);
    const encoder: TextEncoder = new TextEncoder();
    const encodedBytes: Uint8Array = encoder.encode(auditEnvelopeJson);

    const hashBuffer: ArrayBuffer = await window.crypto.subtle.digest(
      "SHA-256",
      encodedBytes as unknown as BufferSource
    );
    const hashArray: number[] = Array.from(new Uint8Array(hashBuffer));
    const recomputedHash: string = hashArray
      .map((b: number): string => b.toString(16).padStart(2, "0"))
      .join("");

    return recomputedHash === storedHash.trim();

  } catch {
    return false;
  }
}

async function generateSecureAuditRecordInternal(context: SignerContext & { signatureData: string }): Promise<SharePointAuditRecord> {
  const { payload, signer, signatureData } = context;
  if (!payload || Object.keys(payload).length === 0) throw new Error("Payload cannot be empty.");
  if (!signer || signer.trim() === "") throw new Error("Signer identifier is required for sealing.");

  const signatureTimestamp: string = new Date().toISOString();
  const sortedPayloadString: string = JSON.stringify(payload, Object.keys(payload).sort());
  const auditEnvelope: AuditEnvelopeRecord = {
    payload: JSON.parse(sortedPayloadString) as Record<string, unknown>,
    signer: signer.toLowerCase().trim(),
    timestamp: signatureTimestamp.trim()
  };
  const auditEnvelopeJson: string = JSON.stringify(auditEnvelope);
  const encoder: TextEncoder = new TextEncoder();
  const encodedBytes: Uint8Array = encoder.encode(auditEnvelopeJson);
  const hashBuffer: ArrayBuffer = await window.crypto.subtle.digest("SHA-256", encodedBytes as unknown as BufferSource);
  const hashArray: number[] = Array.from(new Uint8Array(hashBuffer));
  const signatureHash: string = hashArray.map((b: number): string => b.toString(16).padStart(2, "0")).join("");

  return {
    signatureHash,
    signatureData,
    signatureTimestamp,
    verificationItemId: context.itemID
  };
}