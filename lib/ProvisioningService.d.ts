/**
 * @file EnsurePendingVerificationsList.ts
 *
 * Utility module responsible for idempotent provisioning and configuration
 * of the SharePoint list used for Two-Factor Authentication (2FA) / OTP queueing.
 * Designed to be executed during admin installation or setup checks.
 *
 * @module EnsurePendingVerificationsList
 */
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import { WebPartContext } from "@microsoft/sp-webpart-base";
export declare const DEFAULT_VERIFICATION_LIST_NAME = "PendingVerifications";
export declare const DEFAULT_VERIFICATION_LIST_DESC = "Stores temporary verification passcode records for digital signature workflows.";
export interface IEnsurePendingVerificationsListProps {
    context: WebPartContext;
    listTitle?: string;
    listDescription?: string;
}
/**
 * Idempotently provisions the 2FA verification list, adds required fields,
 * and enforces security. Intended to be run during admin setup/installation.
 */
export declare const ensurePendingVerificationsList: (props: IEnsurePendingVerificationsListProps) => Promise<void>;
