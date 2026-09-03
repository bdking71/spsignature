"use strict";
/**
 * @file SignatureDisplay.tsx
 *
 * React component that renders a signature verification display card
 * showing the signature image, signer information, validation status,
 * timestamp, and cryptographic hash.
 *
 * @module SignatureDisplay
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.SignatureDisplay = void 0;
const react_1 = __importDefault(require("react"));
/**
 * Formats an ISO-8601 timestamp into a localized human-readable date and time string.
 *
 * @param isoTimestamp - ISO-8601 formatted timestamp string.
 * @returns Formatted date/time string in local timezone (e.g., "8/31/2026, 7:55:06 PM").
 */
function formatLocalDateTime(isoTimestamp) {
    try {
        const date = new Date(isoTimestamp);
        return date.toLocaleString();
    }
    catch (_a) {
        return isoTimestamp;
    }
}
/**
 * Truncates a long hash string to a specified length with ellipsis.
 *
 * @param hash - Full hash string.
 * @param length - Number of characters to display before ellipsis (default: 40).
 * @returns Truncated hash string with "..." suffix.
 */
function truncateHash(hash, length = 40) {
    if (hash.length <= length)
        return hash;
    return hash.substring(0, length) + "...";
}
/**
 * Extracts the signer's display name or email from the audit record.
 * Falls back to "Unknown Signer" if not available.
 *
 * @param auditRecord - The audit record containing signature data.
 * @returns The signer identifier or a default fallback string.
 */
function getSignerDisplay(auditRecord) {
    // The audit record doesn't store signer name by design (for security),
    // so consumers must pass it separately. This is a fallback.
    return "Signature";
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
const SignatureDisplay = ({ auditRecord, isValid, className, style, }) => {
    const signatureImageSrc = auditRecord.signatureData;
    const formattedDateTime = formatLocalDateTime(auditRecord.signatureTimestamp);
    const truncatedHash = truncateHash(auditRecord.signatureHash);
    const validationStatusColor = isValid ? "#107c10" : "#d13438";
    const validationStatusLabel = isValid
        ? "Signature is Valid"
        : "Signature is Invalid";
    return (react_1.default.createElement("div", { className: className, style: Object.assign({ border: "1px solid #e0e0e0", borderRadius: "8px", padding: "16px", backgroundColor: "#ffffff", fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif" }, style) },
        react_1.default.createElement("div", { style: {
                marginBottom: "16px",
                textAlign: "center",
                padding: "12px",
                backgroundColor: "#fafafa",
                borderRadius: "6px",
                border: "1px solid #e0e0e0",
            } },
            react_1.default.createElement("img", { src: signatureImageSrc, alt: "Digital Signature", style: {
                    maxWidth: "100%",
                    height: "auto",
                    minHeight: "80px",
                    maxHeight: "200px",
                    display: "block",
                    margin: "0 auto",
                    objectFit: "contain",
                } })),
        react_1.default.createElement("div", { style: {
                marginBottom: "12px",
                fontSize: "16px",
                fontWeight: "600",
                color: "#323130",
            } }, getSignerDisplay(auditRecord)),
        react_1.default.createElement("div", { style: {
                marginBottom: "12px",
                fontSize: "14px",
                fontWeight: "600",
                color: validationStatusColor,
            } }, validationStatusLabel),
        react_1.default.createElement("div", { style: {
                marginBottom: "12px",
                fontSize: "14px",
                color: "#323130",
                display: "flex",
                alignItems: "center",
                gap: "8px",
            } },
            react_1.default.createElement("span", null, "\uD83D\uDCC5"),
            formattedDateTime),
        react_1.default.createElement("div", { style: {
                fontSize: "11px",
                color: "#999999",
                backgroundColor: "#f5f5f5",
                padding: "8px 12px",
                borderRadius: "4px",
                wordBreak: "break-all",
                fontFamily: "'Courier New', monospace",
                border: "1px solid #e0e0e0",
                maxHeight: "60px",
                overflowY: "auto",
            }, title: auditRecord.signatureHash },
            react_1.default.createElement("strong", { style: { color: "#666666" } }, "Hash: "),
            truncatedHash)));
};
exports.SignatureDisplay = SignatureDisplay;
exports.default = exports.SignatureDisplay;
