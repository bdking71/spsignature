/**
 * @file SecureAuditSignature.ts
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
import { DeliveryChannel } from "./OtpService";
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
     * Defaults to `true` when omitted or set to `undefined`.
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
}
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
 *   requireTFA: true,
 * });
 * ```
 */
export declare function promptAndGenerateSecureAudit(context: SignerContext, modalTitle?: string, warningMessage?: string): Promise<SharePointAuditRecord | undefined>;
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
export declare function getReportableSignature(compressedSignatureData: string): string;
/**
 * Re-computes the SHA-256 hash of the audit envelope constructed from
 * the supplied parameters and compares it to `storedHash`.
 *
 * This allows any consumer to independently verify that a signed
 * record has not been tampered with, without needing access to the
 * original signature image.
 *
 * **Note:** Only top-level payload keys are sorted.  If your payload
 * contains nested objects whose key order may vary, consider using a
 * deep-sort utility before calling this function.
 *
 * @param payloadToVerify         - The payload that was originally signed.
 * @param signer                  - The signer identifier used at sign time.
 * @param timestamp               - The ISO-8601 timestamp captured at sign
 *                                  time.
 * @param _compressedSignatureData - The compressed signature (unused by the
 *                                  hash, retained for API symmetry).
 * @param storedHash              - The SHA-256 hex digest to compare against.
 * @returns `true` if the recomputed hash matches `storedHash`.
 */
export declare function verifySecureAuditRecord(payloadToVerify: Record<string, unknown>, signer: string, timestamp: string, _compressedSignatureData: string, storedHash: string): Promise<boolean>;
