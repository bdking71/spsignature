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

    const extractedId = addedItem?.data?.Id ?? addedItem?.data?.ID;
    const itemId = typeof extractedId === "number" && extractedId > 0 ? extractedId : undefined;

    if (!itemId) {
      throw new Error("Item was created, but failed to retrieve a valid SharePoint Item ID.");
    }

    return {
      success: true,
      itemId,
    };
  } catch (error: unknown) {
    const errorMsg = error instanceof Error ? error.message : String(error);
    return {
      success: false,
      error: errorMsg,
    };
  }
}