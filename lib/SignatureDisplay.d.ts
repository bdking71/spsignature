/**
 * @file SignatureDisplay.tsx
 *
 * React component that renders a signature verification display card
 * showing the signature image, signer information, validation status,
 * timestamp, and cryptographic hash.
 *
 * @module SignatureDisplay
 */
import React from "react";
import { SharePointAuditRecord } from "./TransactionSigner";
/**
 * Props for the SignatureDisplay component.
 */
export interface ISignatureDisplayProps {
    /** The audit record to display containing signature data and metadata. */
    auditRecord: SharePointAuditRecord;
    /** Whether the signature has been cryptographically verified as authentic. */
    isValid: boolean;
    /** Optional custom CSS class name for the root container. */
    className?: string;
    /** Optional inline styles for the root container. */
    style?: React.CSSProperties;
}
/**
 * React component that displays a digital signature with verification status,
 * signer information, timestamp, and hash digest.
 *
 * Renders a structured div containing:
 * - Signature image preview
 * - Validation status badge (green for valid, red for invalid)
 * - Signer name/email
 * - Signature date and time (localized)
 * - SHA-256 hash digest (truncated, light gray)
 *
 * All styling is customizable via `className` and `style` props, allowing
 * consumers to integrate the component into any design system.
 *
 * @param props - Component props.
 * @returns JSX.Element representing the signature display.
 *
 * @example
 * ```tsx
 * import { SignatureDisplay } from "@bdking71/spsignature";
 *
 * const MyComponent: React.FC = () => {
 *   const [auditRecord, setAuditRecord] = React.useState<SharePointAuditRecord | null>(null);
 *   const [isValid, setIsValid] = React.useState(false);
 *
 *   return (
 *     <SignatureDisplay
 *       auditRecord={auditRecord}
 *       isValid={isValid}
 *       className="my-signature-display"
 *       style={{ padding: "20px", border: "1px solid #ccc" }}
 *     />
 *   );
 * };
 * ```
 */
export declare const SignatureDisplay: React.FC<ISignatureDisplayProps>;
export default SignatureDisplay;
