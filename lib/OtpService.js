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
    var _a, _b, _c, _d;
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
        // DEBUG: Log the entire response structure
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const debugItem = addedItem;
        console.log("=== PnPjs Add Item Response ===");
        console.log("Full Response:", debugItem);
        console.log("Response Keys:", Object.keys(debugItem || {}));
        if (debugItem === null || debugItem === void 0 ? void 0 : debugItem.data) {
            console.log("Data Object:", debugItem.data);
            console.log("Data Keys:", Object.keys(debugItem.data));
        }
        // Try multiple extraction paths
        let itemId;
        // Path 1: addedItem.data.Id (most common in PnPjs v3/v4)
        if ((_c = addedItem === null || addedItem === void 0 ? void 0 : addedItem.data) === null || _c === void 0 ? void 0 : _c.Id) {
            itemId = addedItem.data.Id;
            console.log("✓ Found ID in addedItem.data.Id:", itemId);
        }
        // Path 2: addedItem.data.ID (alternative casing)
        else if ((_d = addedItem === null || addedItem === void 0 ? void 0 : addedItem.data) === null || _d === void 0 ? void 0 : _d.ID) {
            itemId = addedItem.data.ID;
            console.log("✓ Found ID in addedItem.data.ID:", itemId);
        }
        // Path 3: addedItem.Id (direct property)
        else if (debugItem === null || debugItem === void 0 ? void 0 : debugItem.Id) {
            itemId = debugItem.Id;
            console.log("✓ Found ID in addedItem.Id:", itemId);
        }
        // Path 4: addedItem.ID (direct property, alternate casing)
        else if (debugItem === null || debugItem === void 0 ? void 0 : debugItem.ID) {
            itemId = debugItem.ID;
            console.log("✓ Found ID in addedItem.ID:", itemId);
        }
        // Path 5: addedItem.id (lowercase)
        else if (debugItem === null || debugItem === void 0 ? void 0 : debugItem.id) {
            itemId = debugItem.id;
            console.log("✓ Found ID in addedItem.id:", itemId);
        }
        // Validate the extracted ID
        if (!itemId || typeof itemId !== "number" || itemId <= 0) {
            console.error("❌ Item ID extraction failed. Extracted value:", itemId);
            throw new Error(`Item created but ID extraction failed. Got: ${JSON.stringify(itemId)}`);
        }
        console.log(`✓ OTP Record created successfully with ID: ${itemId}`);
        return {
            success: true,
            itemId,
        };
    }
    catch (error) {
        const errorMsg = error instanceof Error ? error.message : String(error);
        console.error("❌ dispatchOtpPasscode Error:", errorMsg);
        console.error("Full Error Object:", error);
        return {
            success: false,
            error: errorMsg,
        };
    }
}
exports.dispatchOtpPasscode = dispatchOtpPasscode;
