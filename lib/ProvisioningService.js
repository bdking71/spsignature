"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.ensurePendingVerificationsList = exports.DEFAULT_VERIFICATION_LIST_DESC = exports.DEFAULT_VERIFICATION_LIST_NAME = void 0;
const sp_1 = require("@pnp/sp");
require("@pnp/sp/webs");
require("@pnp/sp/lists");
require("@pnp/sp/fields");
/**
 * Default SharePoint list title used for queuing 2FA verification requests.
 */
exports.DEFAULT_VERIFICATION_LIST_NAME = "PendingVerifications";
/**
 * Default SharePoint list description.
 */
exports.DEFAULT_VERIFICATION_LIST_DESC = "Stores temporary verification passcode records for digital signature workflows.";
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
const ensurePendingVerificationsList = async (props) => {
    if (!props.context) {
        throw new Error("Execution restricted: Valid WebPartContext is required.");
    }
    const listTitle = props.listTitle || exports.DEFAULT_VERIFICATION_LIST_NAME;
    const listDescription = props.listDescription || exports.DEFAULT_VERIFICATION_LIST_DESC;
    const sp = (0, sp_1.spfi)().using((0, sp_1.SPFx)(props.context));
    let list = sp.web.lists.getByTitle(listTitle);
    // ---------------------------------------------------------------------------
    // 1. Ensure the List Exists
    // ---------------------------------------------------------------------------
    try {
        await list();
    }
    catch (_a) {
        // 100 = Generic List template, true = enable content approval/visibility
        await sp.web.lists.add(listTitle, listDescription, 100, true);
        list = sp.web.lists.getByTitle(listTitle);
    }
    // ---------------------------------------------------------------------------
    // 2. Helper: Ensure Field Exists (Idempotent Field Creation)
    // ---------------------------------------------------------------------------
    /**
     * Checks if a field exists by internal name or title; if missing, calls the creation callback.
     *
     * @param fieldName     - Internal name or title of the field.
     * @param createFieldFn - Async callback to execute if the field does not exist.
     */
    const ensureField = async (fieldName, createFieldFn) => {
        try {
            await list.fields.getByInternalNameOrTitle(fieldName)();
        }
        catch (_a) {
            // Field does not exist, provision it
            await createFieldFn();
        }
    };
    // ---------------------------------------------------------------------------
    // 3. Ensure 'Passcode' Field
    // ---------------------------------------------------------------------------
    await ensureField("Passcode", async () => {
        await list.fields.addText("Passcode", {
            Required: true,
            Description: "Temporary 5-digit signature verification passcode.",
        });
    });
    // ---------------------------------------------------------------------------
    // 4. Ensure 'Channel' Field
    // ---------------------------------------------------------------------------
    await ensureField("Channel", async () => {
        await list.fields.addChoice("Channel", {
            Choices: ["email", "teams", "both"],
            Required: true,
            FillInChoice: false,
            Description: "Target notification delivery route (email, teams, or both).",
        });
        const channelField = list.fields.getByInternalNameOrTitle("Channel");
        await channelField.update({ DefaultValue: "email" });
    });
    // ---------------------------------------------------------------------------
    // 5. Enforce Item-Level Security Settings
    // ---------------------------------------------------------------------------
    // ReadSecurity: 2  -> Users can only read items that were created by themselves.
    // WriteSecurity: 2 -> Users can only edit/delete items created by themselves.
    // This prevents end-users from inspecting or hijacking other users' OTP codes.
    await list.update({
        ReadSecurity: 2,
        WriteSecurity: 2,
    });
};
exports.ensurePendingVerificationsList = ensurePendingVerificationsList;
