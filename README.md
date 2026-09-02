# @bdking71/spsignature

@bdking71/spsignature provides a secure, lightweight, and framework-agnostic digital signature module engineered specifically for Microsoft 365 environments. Built to run inside custom SPFx Web Parts and Extension Application Customizers, it delivers tamper-evident signature collection, automated base64 image encoding, and structured payload generation directly integrated with SharePoint Online list infrastructure.

Key Capabilities & Business Value

- Zero Infrastructure Overhead: Operates entirely within client-side M365 contexts—no external APIs, azure functions, or third-party storage backends required.

- Cross-Platform M365 Integration: Designed for seamless deployment across SharePoint Online, Microsoft Teams, and Viva Connections desktop/mobile experiences.

- Integrity & Non-Repudiation: Generates cryptographic SHA-256 hashes bound to user identities, timestamps, and target record IDs to prevent signature tampering.

- Standardized JSON Payload: Encapsulates complete signing metadata inside standard SharePoint Multiline Text columns (Plain Text / JSON formatted) for effortless integration with Power Automate and Power BI.

## Two-Factor Authentication (2FA / TFA) Architecture

To enforce non-repudiation and meet compliance standards, @bdking71/spsignature includes a flexible Two-Factor Authentication engine. Rather than relying on rigid third-party SMS gateways, it delegates code delivery to Power Automate, allowing organizations to route standard 6-digit verification codes via Microsoft Teams, Outlook Email, or both.

![Diagram illustrating the 2FA workflow between SPFx, SharePoint, and Power Automate](./img/tfa.jpg)

## How The 2FA Workflow Operates-
- Code Generation: When a user initiates a signing transaction, the component generates a cryptographically secure, time-sensitive verification code (e.g., 849204) and registers a pending transaction state.

- Power Automate Trigger: The component triggers a light HTTP webhook or writes a record to a dedicated SharePoint verification list.

- Multi-Channel Delivery: The Power Automate flow instantly dispatches the code to the signer via:

- Microsoft Teams: Direct Adaptive Card or Activity Feed notification.

- Outlook Email: High-priority internal notification.

- Validation & Signing: The signer enters the 6-digit code into the @bdking71/spsignature UI component. Upon successful validation, the final signed JSON payload is generated and stored in the target record.

## Signature Data Schema

The module returns a single, structured JSON object designed to be stored directly inside a SharePoint Multiple lines of text field (Plain Text format):

```JSON
{
  "signatureHash": "e3b0c4429 ...",
  "signatureData": "data:image/png;base64,iVB....",
  "signatureTimestamp": "2026-08-31T19:55:06.121Z",
  "verificationItemId": 1
}
```


## Toolchain & Compatibility Matrix

This project requires a precise local environment build. Strict adherence to the versioning matrix is mandatory to prevent toolchain compilation errors.

* **SPFx Version:** `v1.22.1`
* **Node.js:** `v18.17.1` (Recommended use of `nvm` or `nvs`)
* **Gulp CLI:** `v4.x`
* **Primary UI Framework:** React `v17.0.1` / Fluent UI `v8.x`

## Quick Start & Local Development

Execute these commands in sequence to establish your local development runtime:

## Integration & Usage Guide

### Provision Infrastructure
Call `ensurePendingVerificationsList` during the WebPart lifecycle (e.g., inside `onInit()`) to automatically ensure the required `PendingVerifications` SharePoint list and columns exist on the site collection[cite: 2].

```typescript
import { ensurePendingVerificationsList } from "@bdking71/spsignature";

public async onInit(): Promise<void> {
  await super.onInit();

  // Ensures 'PendingVerifications' list with 'Passcode' and 'Channel' fields exists
  await ensurePendingVerificationsList({ context: this.context });
}
```

### Launch Signature Modal & Capture Audit

Execute promptAndGenerateSecureAudit within your component or event handler. This renders an enterprise modal overlay that handles draw, upload, or cached signature selection, triggers 2FA code delivery, validates the 5-digit passcode, and outputs a cryptographic audit payload.

```TypeScript
import {
  promptAndGenerateSecureAudit,
  SharePointAuditRecord
} from "@bdking71/spsignature";

private handleSignAction = async (): Promise<void> => {
  const auditRecord: SharePointAuditRecord | undefined = await promptAndGenerateSecureAudit(
    {
      itemID: 1042, // Record/Item ID being signed
      signer: this.context.pageContext.user.email,
      spContext: this.context,
      channel: "teams", // Code delivery route: "teams" | "email" | "both"
      requireTFA: true,
      payload: {
        DocumentTitle: "Purchase Requisition #4402",
        TotalCost: "$12,450.00",
        ApprovedBy: this.context.pageContext.user.displayName
      }
    },
    "Approve Purchase Requisition",
    "Please review your entry and apply your signature to finalize this record."
  );

  if (auditRecord) {
    // Save auditRecord JSON payload directly to your SharePoint multiline text column
    console.log("Signature Hash:", auditRecord.signatureHash);
    console.log("Compressed Image Data:", auditRecord.signatureData);
    console.log("Timestamp:", auditRecord.signatureTimestamp);
    console.log("Verification Log ID:", auditRecord.verificationItemId);
  }
};
```

### Verification & Image Reconstruction

To verify non-repudiation or render the stored LZW-compressed signature image back into a view or report, use the built-in helper utilities.

```typescript
import {
  verifySecureAuditRecord,
  getReportableSignature
} from "@bdking71/spsignature";

// 1. Re-calculate SHA-256 hash to verify record integrity
const isValid: boolean = await verifySecureAuditRecord(
  payload,
  signerEmail,
  timestamp,
  compressedSignatureData,
  storedHash
);

// 2. Decompress LZW signature string into a standard PNG base64 Data-URI
const renderableImageSrc: string = getReportableSignature(compressedSignatureData);
```