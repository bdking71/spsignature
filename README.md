# @bdking71/spsignature

[![GitHub Repository](https://img.shields.io/badge/GitHub-Repository-blue?logo=github)](https://github.com/bdking71/spsignature)
[![npm version](https://img.shields.io/npm/v/@bdking71/spsignature.svg)](https://www.npmjs.com/package/@bdking71/spsignature)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

> ⚠️ **NOT READY FOR PRODUCTION USE!** This module is currently under active development. Please do not use it in production environments at this time.

![Signature Screen Screenshot](/img/SignatureScreen.png)

`@bdking71/spsignature` provides a secure, lightweight, and framework-agnostic digital signature module engineered specifically for Microsoft 365 environments. Built to run inside custom SPFx Web Parts and Extension Application Customizers, it delivers tamper-evident signature collection, automated base64 image encoding, and structured payload generation directly integrated with SharePoint Online list infrastructure.

---

## Source Code & Repository

* **GitHub Repository:** https://github.com/bdking71/spsignature
* **Issues & Feedback:** https://github.com/bdking71/spsignature/issues
* **NPM Package:** https://www.npmjs.com/package/@bdking71/spsignature

---

## Key Capabilities & Business Value

- **Zero Infrastructure Overhead**: Operates entirely within client-side M365 contexts — no external APIs, Azure Functions, or third-party storage backends required.
- **Cross-Platform M365 Integration**: Designed for seamless deployment across SharePoint Online, Microsoft Teams, and Viva Connections desktop/mobile experiences.
- **Integrity & Non-Repudiation**: Generates cryptographic SHA-256 hashes bound to user identities, timestamps, and target record IDs to prevent signature tampering.
- **Standardized JSON Payload**: Encapsulates complete signing metadata inside standard SharePoint Multiline Text columns for effortless integration with Power Automate and Power BI.
- **LZW Compression**: Automatically compresses signature image payloads to minimize SharePoint storage footprint.
- **Signature Caching**: Optional local device caching so users don't have to re-draw their signature on every transaction.
- **React Component Display**: Built-in `SignatureDisplay` React component for displaying verified signatures with full styling control.

---

## Two-Factor Authentication (2FA) Architecture

To enforce non-repudiation and meet compliance standards, `@bdking71/spsignature` includes a flexible Two-Factor Authentication engine. Rather than relying on rigid third-party SMS gateways, it delegates code delivery to **Power Automate**, allowing organizations to route standard **5-digit** verification codes via Microsoft Teams, Outlook Email, or both.

### How the 2FA Workflow Operates
![Diagram illustrating the 2FA workflow between SPFx, SharePoint, and Power Automate](/img/TFA.jpg)

1. **Code Generation**: When a user initiates a signing transaction, the component generates a cryptographically secure 5-digit verification code (e.g., `84920`) using the Web Crypto API.
2. **SharePoint Queue**: The component writes a record to the dedicated `PendingVerifications` SharePoint list, secured with item-level permissions so users can only see their own codes.
3. **Power Automate Trigger**: A Power Automate flow monitors the list and instantly dispatches the code to the signer via:
   - **Microsoft Teams**: Direct Adaptive Card or Activity Feed notification.
   - **Outlook Email**: High-priority internal notification.
4. **Validation & Signing**: The signer enters the 5-digit code into the `@bdking71/spsignature` UI component. Upon successful validation, the final signed JSON payload is generated and stored in the target record.

---

## Signature Data Schema

The module returns a single, structured JSON object designed to be stored directly inside a SharePoint **Multiple lines of text** field:

```
{
  "signatureHash": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
  "signatureData": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAA...",
  "signatureTimestamp": "2026-08-31T19:55:06.121Z",
  "verificationItemId": 42
}
```

| Field | Type | Description |
| :--- | :--- | :--- |
| `signatureHash` | `string` | SHA-256 hex digest of the canonical audit envelope (payload + signer + timestamp). |
| `signatureData` | `string` | LZW-compressed data-URI of the signature image (PNG). |
| `signatureTimestamp` | `string` | ISO-8601 timestamp captured at the moment of signing. |
| `verificationItemId` | `number` | ID of the corresponding record in the `PendingVerifications` list. |

---

## Toolchain & Compatibility Matrix

This project requires a precise local environment build. Strict adherence to the versioning matrix is mandatory to prevent toolchain compilation errors.

| Component | Version |
| :--- | :--- |
| **SPFx** | `v1.22.1` |
| **Node.js** | `v18.17.1` (Recommended use of `nvm` or `nvs`) |
| **Gulp CLI** | `v4.x` |
| **React** | `v17.0.1` |
| **Fluent UI** | `v8.x` |
| **TypeScript** | `v4.7+` |

---

## Installation

Install the package into your SPFx solution:

```
npm install @bdking71/spsignature --save
```

### Peer Dependencies

Ensure the following peer dependencies are installed in your SPFx project:

```
npm install @pnp/sp @microsoft/sp-webpart-base react --save
```

---

## Integration & Usage Guide

### Step 1: Provision SharePoint Infrastructure

Call `ensurePendingVerificationsList` during the web part lifecycle (inside `onInit()`) to automatically ensure the `PendingVerifications` SharePoint list, columns, and item-level security settings exist on the current site collection.

Note: This method is fully **idempotent**—it can be safely called on every web part load without failing or duplicating columns.

```
import { ensurePendingVerificationsList } from "@bdking71/spsignature";

export default class MySignatureWebPart extends BaseClientSideWebPart<IMySignatureWebPartProps> {

  public async onInit(): Promise<void> {
    await super.onInit();

    // Provisions the 'PendingVerifications' list with 'Passcode' and 'Channel' columns
    await ensurePendingVerificationsList({ context: this.context });
  }
}
```

### Step 2: Configure Your Power Automate Flow

Create a Power Automate flow that triggers on **"When an item is created"** for the `PendingVerifications` list. The flow should:

1. Read the `Title` (signer email), `Passcode`, and `Channel` columns.
2. Route notification based on the `Channel` value:
   - `"email"` → Send an Outlook email with the passcode.
   - `"teams"` → Post an Adaptive Card to the user in Microsoft Teams.
   - `"both"` → Dispatch through both channels simultaneously.
3. Optionally delete the item after dispatch (or use a scheduled cleanup flow to purge expired codes).

### Step 3: Launch Signature Modal & Capture Audit

Execute `promptAndGenerateSecureAudit` within your component or event handler. This renders an enterprise modal overlay that:
- Handles drawn, uploaded, or cached signature selection.
- Triggers 2FA code delivery through the configured channel (when enabled).
- Validates the 5-digit passcode (if TFA is required).
- Outputs a cryptographic audit payload.

```
import {
  promptAndGenerateSecureAudit,
  SharePointAuditRecord
} from "@bdking71/spsignature";

private handleSignAction = async (): Promise<void> => {
  const auditRecord: SharePointAuditRecord | undefined = await promptAndGenerateSecureAudit(
    {
      itemID: 1042,
      signer: this.context.pageContext.user.email,
      spContext: this.context,
      channel: "teams",
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
    console.log("Signature Hash:", auditRecord.signatureHash);
    console.log("Compressed Image Data:", auditRecord.signatureData);
    console.log("Timestamp:", auditRecord.signatureTimestamp);
    console.log("Verification Log ID:", auditRecord.verificationItemId);

    await sp.web.lists.getByTitle("Requisitions").items.getById(1042).update({
      SignatureAudit: JSON.stringify(auditRecord)
    });
  }
};
```

### Step 4: Verification & Image Reconstruction

To verify non-repudiation or render the stored signature back into a view or report, use the built-in helper utilities.

```
import {
  verifySecureAuditRecord,
  getReportableSignature
} from "@bdking71/spsignature";

// 1. Re-calculate SHA-256 hash to verify record integrity
const isValid: boolean = await verifySecureAuditRecord(
  auditRecord,
  signerEmail,
  payload
);

if (isValid) {
  console.log("✓ Signature is authentic and untampered.");
} else {
  console.warn("✗ Signature verification failed - record may be tampered.");
}

// 2. Decompress LZW signature string into a viewable PNG Data-URI
const renderableImageSrc: string = getReportableSignature(auditRecord.signatureData);
```

### Step 5: Display Signature with React Component

Use the built-in `SignatureDisplay` React component to render a verified signature with customizable styling:

```
import React from "react";
import {
  SignatureDisplay,
  SharePointAuditRecord,
  verifySecureAuditRecord
} from "@bdking71/spsignature";

export const MySignatureViewer: React.FC = () => {
  const [auditRecord, setAuditRecord] = React.useState<SharePointAuditRecord | null>(null);
  const [isValid, setIsValid] = React.useState(false);

  React.useEffect(() => {
    const verifySignature = async () => {
      if (!auditRecord) return;

      const valid = await verifySecureAuditRecord(
        auditRecord,
        "user@example.com",
        { amount: 1500, vendor: "Contoso" }
      );
      setIsValid(valid);
    };

    verifySignature();
  }, [auditRecord]);

  if (!auditRecord) {
    return <div>No signature to display.</div>;
  }

  return (
    <div>
      <h1>Purchase Requisition #4402</h1>

      <SignatureDisplay
        auditRecord={auditRecord}
        isValid={isValid}
        style={{
          padding: "24px",
          border: "2px solid #0078d4",
          backgroundColor: "#f0f7ff",
          marginTop: "20px",
          borderRadius: "12px",
        }}
      />

      <SignatureDisplay
        auditRecord={auditRecord}
        isValid={isValid}
        className="my-signature-card"
      />
    </div>
  );
};
```

#### SignatureDisplay Component Props

| Prop | Type | Required | Description |
| :--- | :--- | :--- | :--- |
| `auditRecord` | `SharePointAuditRecord` | Yes | The signature audit record returned from `promptAndGenerateSecureAudit()` |
| `isValid` | `boolean` | Yes | Result from `verifySecureAuditRecord()` — displayed in green or red |
| `className` | `string` | No | CSS class name for custom styling |
| `style` | `React.CSSProperties` | No | Inline CSS styles for the root container |

The SignatureDisplay component renders:
- Signature image preview
- Signer name/email
- Validation status (green for valid, red for invalid)
- ISO-8601 timestamp converted to local timezone
- SHA-256 hash digest (truncated, light gray font)

---

## API Reference

### Exported Functions

| Function | Purpose |
| :--- | :--- |
| `ensurePendingVerificationsList(props)` | Idempotently provisions the `PendingVerifications` SharePoint list, columns, and item-level security. |
| `promptAndGenerateSecureAudit(context, title?, msg?)` | Launches the signature modal, dispatches the OTP, validates entry, and returns a `SharePointAuditRecord`. |
| `verifySecureAuditRecord(auditRecord, signer, payload)` | Re-computes the SHA-256 hash to verify audit-record integrity. |
| `getReportableSignature(compressedSignatureData)` | Decompresses an LZW-compressed signature back into a viewable `data:image/png` URI. |

### Exported Types

| Type | Description |
| :--- | :--- |
| `SharePointAuditRecord` | Structured JSON audit payload returned after successful signing. |
| `SignerContext` | Input parameters for the signature workflow. |
| `AuditEnvelopeRecord` | Canonical envelope structure used for hashing (payload + signer + timestamp). |
| `DeliveryChannel` | `"email" \| "teams" \| "both"` — Supported OTP delivery channels. |
| `ISignatureDisplayProps` | Props interface for the `SignatureDisplay` React component. |

### Exported React Components

| Component | Purpose |
| :--- | :--- |
| `SignatureDisplay` | Renders a verified signature with signer info, validation status, timestamp, and hash digest. Fully customizable via props. |

---

## Security Considerations

- **Client-Side Passcode Generation**: The 5-digit OTP is generated and validated in the browser using the Web Crypto API. This is well-suited for **enterprise workflow enforcement** (approvals, sign-offs, SOP compliance) but is not intended to replace hardened server-side MFA solutions for high-risk financial transactions.
- **Item-Level Security**: The `PendingVerifications` list is automatically configured with `ReadSecurity = 2` and `WriteSecurity = 2`, meaning users can only view and edit records they created. This prevents users from inspecting other users' active passcodes.
- **SHA-256 Non-Repudiation**: The signature hash binds the payload, signer identity, and timestamp together. Any modification to the stored record can be detected through `verifySecureAuditRecord()`.
- **Signature Caching**: Cached signatures are stored in browser `localStorage` after LZW compression. Consumers should clearly communicate this caching behavior to end users to meet compliance and privacy requirements (GDPR, CCPA, etc.).

---

## Troubleshooting

| Issue | Resolution |
| :--- | :--- |
| Modal never appears | Ensure `spContext` is a valid `WebPartContext` and the DOM is fully loaded. |
| OTP never arrives | Verify the Power Automate flow is active and listening on the correct list. Check the flow run history. |
| SHA-256 verification fails on nested objects | Note: Only top-level payload keys are sorted. Deeply nested objects should have consistent key ordering. |
| Signature not caching | Verify browser `localStorage` isn't in private/incognito mode or blocked by browser policy. |
| SignatureDisplay not rendering | Ensure React is installed and `auditRecord` is properly populated with valid audit data. |

---

## Contributing

Contributions, issues, and feature requests are welcome!

1. Fork the Project: https://github.com/bdking71/spsignature
2. Create your Feature Branch (git checkout -b feature/AmazingFeature)
3. Commit your Changes (git commit -m 'Add some AmazingFeature')
4. Push to the Branch (git push origin feature/AmazingFeature)
5. Open a Pull Request

---

## License

MIT License

Copyright (c) 2025 bdking71

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.