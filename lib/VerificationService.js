"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.createPendingVerification = void 0;
const sp_1 = require("@pnp/sp");
require("@pnp/sp/webs");
require("@pnp/sp/lists");
require("@pnp/sp/items");
async function createPendingVerification(context, request) {
    var _a, _b, _c, _d, _e;
    try {
        if (!context) {
            throw new Error("Execution restricted: Valid WebPartContext is required.");
        }
        const sp = (0, sp_1.spfi)().using((0, sp_1.SPFx)(context));
        const LIST_NAME = "PendingVerifications";
        const addedItem = await sp.web.lists.getByTitle(LIST_NAME).items.add({
            Title: request.title,
            Passcode: request.passcode,
            Channel: request.channel,
        });
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const itemId = (_e = (_d = (_b = (_a = addedItem === null || addedItem === void 0 ? void 0 : addedItem.data) === null || _a === void 0 ? void 0 : _a.Id) !== null && _b !== void 0 ? _b : (_c = addedItem === null || addedItem === void 0 ? void 0 : addedItem.data) === null || _c === void 0 ? void 0 : _c.ID) !== null && _d !== void 0 ? _d : addedItem === null || addedItem === void 0 ? void 0 : addedItem.Id) !== null && _e !== void 0 ? _e : 0;
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
exports.createPendingVerification = createPendingVerification;
