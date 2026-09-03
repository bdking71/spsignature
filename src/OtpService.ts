/**
 * @file OtpService.ts
 * Provides methods for dispatching One-Time Passcodes (OTP) for
 * Two-Factor Authentication via SharePoint and external channels.
 */

import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export type DeliveryChannel = "email" | "teams" | "both";

export interface IOtpDispatchRequest {
  title: string;
  passcode: string;
  channel: DeliveryChannel;
}

export interface IOtpDispatchResult {
  success: boolean;
  itemId?: number;
  error?: string;
}

const DEFAULT_OTP_LIST_NAME = "PendingVerifications";

/**
 * Queues an OTP dispatch request in SharePoint to notify the user via Email/Teams.
 */
export async function dispatchOtpPasscode(
  context: WebPartContext,
  request: IOtpDispatchRequest,
  listTitle: string = DEFAULT_OTP_LIST_NAME
): Promise<IOtpDispatchResult> {
  try {
    if (!context) {
      throw new Error("Execution restricted: Valid WebPartContext is required.");
    }

    if (!request || !request.title?.trim()) {
      throw new Error("Invalid request: Signer/Title identifier is required.");
    }

    if (!request.passcode?.trim()) {
      throw new Error("Invalid request: Passcode must be provided.");
    }

    const sp: SPFI = spfi().using(SPFx(context));

    const addedItem = await sp.web.lists.getByTitle(listTitle).items.add({
      Title: request.title.trim(),
      Passcode: request.passcode.trim(),
      Channel: request.channel || "email",
    });

    // DEBUG: Log the entire response structure
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const debugItem = addedItem as any;
    console.log("=== PnPjs Add Item Response ===");
    console.log("Full Response:", debugItem);
    console.log("Response Keys:", Object.keys(debugItem || {}));
    if (debugItem?.data) {
      console.log("Data Object:", debugItem.data);
      console.log("Data Keys:", Object.keys(debugItem.data));
    }

    // Try multiple extraction paths
    let itemId: number | undefined;

    // Path 1: addedItem.data.Id (most common in PnPjs v3/v4)
    if (addedItem?.data?.Id) {
      itemId = addedItem.data.Id;
      console.log("✓ Found ID in addedItem.data.Id:", itemId);
    }
    // Path 2: addedItem.data.ID (alternative casing)
    else if (addedItem?.data?.ID) {
      itemId = addedItem.data.ID;
      console.log("✓ Found ID in addedItem.data.ID:", itemId);
    }
    // Path 3: addedItem.Id (direct property)
    else if (debugItem?.Id) {
      itemId = debugItem.Id as number;
      console.log("✓ Found ID in addedItem.Id:", itemId);
    }
    // Path 4: addedItem.ID (direct property, alternate casing)
    else if (debugItem?.ID) {
      itemId = debugItem.ID as number;
      console.log("✓ Found ID in addedItem.ID:", itemId);
    }
    // Path 5: addedItem.id (lowercase)
    else if (debugItem?.id) {
      itemId = debugItem.id as number;
      console.log("✓ Found ID in addedItem.id:", itemId);
    }

    // Validate the extracted ID
    if (!itemId || typeof itemId !== "number" || itemId <= 0) {
      console.error("❌ Item ID extraction failed. Extracted value:", itemId);
      throw new Error(
        `Item created but ID extraction failed. Got: ${JSON.stringify(itemId)}`
      );
    }

    console.log(`✓ OTP Record created successfully with ID: ${itemId}`);
    return {
      success: true,
      itemId,
    };
  } catch (error: unknown) {
    const errorMsg = error instanceof Error ? error.message : String(error);
    console.error("❌ dispatchOtpPasscode Error:", errorMsg);
    console.error("Full Error Object:", error);
    return {
      success: false,
      error: errorMsg,
    };
  }
}