/**
 * @file EnsurePendingVerificationsList.ts
 *
 * Utility module responsible for idempotent provisioning and configuration
 * of the SharePoint list used for Two-Factor Authentication (2FA) / OTP queueing.
 *
 * This utility ensures:
 *   1. The target SharePoint list exists (creates it if missing).
 *   2. The required custom schema fields ('Passcode' and 'Channel') are provisioned.
 *   3. Item-level security settings are hardened (ReadSecurity=2, WriteSecurity=2)
 *      so non-admin users cannot read or tamper with other users' verification codes.
 *
 * @module EnsurePendingVerificationsList
 */
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import { WebPartContext } from "@microsoft/sp-webpart-base";
/**
 * Default SharePoint list title used for queuing 2FA verification requests.
 */
export declare const DEFAULT_VERIFICATION_LIST_NAME = "PendingVerifications";
/**
 * Default SharePoint list description.
 */
export declare const DEFAULT_VERIFICATION_LIST_DESC = "Stores temporary verification passcode records for digital signature workflows.";
/**
 * Configuration properties for provisioning the pending verifications list.
 */
export interface IEnsurePendingVerificationsListProps {
    /** The SPFx WebPartContext used to initialize the PnPjs client. */
    context: WebPartContext;
    /** Optional custom title for the SharePoint list (defaults to `"PendingVerifications"`). */
    listTitle?: string;
    /** Optional custom description for the list. */
    listDescription?: string;
}
/**
 * Ensures that the SharePoint list for pending verification codes exists,
 * has the required fields (`Passcode` and `Channel`), and enforces item-level
 * read/write security.
 *
 * This function is fully **idempotent**: it can be safely executed multiple times
 * without failing or duplicating columns.
 *
 * @param props - Configuration properties containing context and optional list overrides.
 * @returns A promise that resolves when the list schema and security are verified.
 *
 * @throws {Error} If `context` is invalid or if permissions prevent list provisioning.
 *
 * @example
 * ```ts
 * await ensurePendingVerificationsList({
 *   context: this.context
 * });
 * ```
 */
export declare const ensurePendingVerificationsList: (props: IEnsurePendingVerificationsListProps) => Promise<void>;
