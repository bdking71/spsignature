"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ensurePendingVerificationsList = void 0;
const sp_1 = require("@pnp/sp");
require("@pnp/sp/webs");
require("@pnp/sp/lists");
require("@pnp/sp/fields");
require("@pnp/sp/views");
const ensurePendingVerificationsList = async (props) => {
    if (!props.context) {
        throw new Error("Execution restricted: Valid WebPartContext is required.");
    }
    const sp = (0, sp_1.spfi)().using((0, sp_1.SPFx)(props.context));
    const LIST_NAME = "PendingVerifications";
    const LIST_DESCRIPTION = "Stores temporary PIN verification codes for document signatures.";
    const list = sp.web.lists.getByTitle(LIST_NAME);
    let listCreated = false;
    try {
        await list();
    }
    catch (_a) {
        await sp.web.lists.add(LIST_NAME, LIST_DESCRIPTION, 100, true);
        listCreated = true;
    }
    if (!listCreated) {
        return;
    }
    const ensureField = async (fieldName, createFieldFn) => {
        try {
            await createFieldFn();
        }
        catch (creationErr) {
            throw creationErr;
        }
    };
    await ensureField("Passcode", async () => {
        await list.fields.addText("Passcode", {
            Required: true,
            Description: "Temporary 6-digit signature verification code."
        });
    });
    await ensureField("Channel", async () => {
        await list.fields.addChoice("Channel", {
            Choices: ["email", "teams", "both"],
            Required: true,
            FillInChoice: false,
            Description: "Target notification delivery route."
        });
        const channelField = list.fields.getByInternalNameOrTitle("Channel");
        await channelField.update({ DefaultValue: "teams" });
    });
    try {
        await list.update({
            ReadSecurity: 2,
            WriteSecurity: 2
        });
    }
    catch (secErr) {
        throw secErr;
    }
};
exports.ensurePendingVerificationsList = ensurePendingVerificationsList;
