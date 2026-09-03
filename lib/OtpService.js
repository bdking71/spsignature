"use strict";
/**
 * @file OtpService.ts
 * Provides methods for dispatching One-Time Passcodes (OTP) for
 * Two-Factor Authentication via SharePoint and external channels.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.dispatchOtpPasscode = void 0;
const sp_1 = require("@pnp/sp");
require("@pnp/sp/webs");
require("@pnp/sp/lists");
require("@pnp/sp/items");
const DEFAULT_OTP_LIST_NAME = "PendingVerifications";
/**
 * Queues an OTP dispatch request in SharePoint to notify the user via Email/Teams.
 */
async function dispatchOtpPasscode(context, request, listTitle = DEFAULT_OTP_LIST_NAME) {
    var _a, _b, _c, _d, _e;
    try {
        if (!context) {
            throw new Error("Execution restricted: Valid WebPartContext is required.");
        }
        if (!request || !((_a = request.title) === null || _a === void 0 ? void 0 : _a.trim())) {
            throw new Error("Invalid request: Signer/Title identifier is required.");
        }
        if (!((_b = request.passcode) === null || _b === void 0 ? void 0 : _b.trim())) {
            throw new Error("Invalid request: Passcode must be provided.");
        }
        const sp = (0, sp_1.spfi)().using((0, sp_1.SPFx)(context));
        const addedItem = await sp.web.lists.getByTitle(listTitle).items.add({
            Title: request.title.trim(),
            Passcode: request.passcode.trim(),
            Channel: request.channel || "email",
        });
        const extractedId = (_d = (_c = addedItem === null || addedItem === void 0 ? void 0 : addedItem.data) === null || _c === void 0 ? void 0 : _c.Id) !== null && _d !== void 0 ? _d : (_e = addedItem === null || addedItem === void 0 ? void 0 : addedItem.data) === null || _e === void 0 ? void 0 : _e.ID;
        const itemId = typeof extractedId === "number" && extractedId > 0 ? extractedId : undefined;
        if (!itemId) {
            throw new Error("Item was created, but failed to retrieve a valid SharePoint Item ID.");
        }
        return {
            success: true,
            itemId,
        };
    }
    catch (error) {
        const errorMsg = error instanceof Error ? error.message : String(error);
        return {
            success: false,
            error: errorMsg,
        };
    }
}
exports.dispatchOtpPasscode = dispatchOtpPasscode;
