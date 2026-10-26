"use strict";
/**
 * @file EnsurePendingVerificationsList.ts
 *
 * Utility module responsible for idempotent provisioning and configuration
 * of the SharePoint list used for Two-Factor Authentication (2FA) / OTP queueing.
 * Designed to be executed during admin installation or setup checks.
 *
 * @module EnsurePendingVerificationsList
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.ensurePendingVerificationsList = exports.DEFAULT_VERIFICATION_LIST_DESC = exports.DEFAULT_VERIFICATION_LIST_NAME = void 0;
const sp_1 = require("@pnp/sp");
require("@pnp/sp/webs");
require("@pnp/sp/lists");
require("@pnp/sp/fields");
exports.DEFAULT_VERIFICATION_LIST_NAME = "PendingVerifications";
exports.DEFAULT_VERIFICATION_LIST_DESC = "Stores temporary verification passcode records for digital signature workflows.";
/**
 * Idempotently provisions the 2FA verification list, adds required fields,
 * and enforces security. Intended to be run during admin setup/installation.
 */
const ensurePendingVerificationsList = async (props) => {
    if (!props.context) {
        throw new Error("Execution restricted: Valid WebPartContext is required.");
    }
    const listTitle = props.listTitle || exports.DEFAULT_VERIFICATION_LIST_NAME;
    const listDescription = props.listDescription || exports.DEFAULT_VERIFICATION_LIST_DESC;
    const sp = (0, sp_1.spfi)().using((0, sp_1.SPFx)(props.context));
    let list = sp.web.lists.getByTitle(listTitle);
    // 1. Ensure the List Exists (Create if missing)
    try {
        await list();
    }
    catch (_a) {
        // 100 = Generic List template
        await sp.web.lists.add(listTitle, listDescription, 100, true);
        list = sp.web.lists.getByTitle(listTitle);
    }
    // 2. Helper: Ensure Field Exists
    const ensureField = async (fieldName, createFieldFn) => {
        try {
            await list.fields.getByInternalNameOrTitle(fieldName)();
        }
        catch (_a) {
            await createFieldFn();
        }
    };
    // 3. Ensure 'Passcode' Field
    await ensureField("Passcode", async () => {
        await list.fields.addText("Passcode", {
            Required: true,
            Description: "Temporary 5-digit signature verification passcode.",
        });
    });
    // 4. Ensure 'Channel' Field
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
    // 5. Enforce Item-Level Security Settings (Admin-configured once)
    await list.update({
        ReadSecurity: 2,
        WriteSecurity: 2,
    });
};
exports.ensurePendingVerificationsList = ensurePendingVerificationsList;
