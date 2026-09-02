import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export type DeliveryChannel = "email" | "teams" | "both";

export interface IVerificationRequest {
  title: string;
  passcode: string;
  channel: DeliveryChannel;
}

export interface IVerificationResult {
  success: boolean;
  itemId?: number;
  error?: string;
}

export async function createPendingVerification(
  context: WebPartContext,
  request: IVerificationRequest
): Promise<IVerificationResult> {
  try {
    if (!context) {
      throw new Error("Execution restricted: Valid WebPartContext is required.");
    }

    const sp: SPFI = spfi().using(SPFx(context));
    const LIST_NAME = "PendingVerifications";

    const addedItem = await sp.web.lists.getByTitle(LIST_NAME).items.add({
      Title: request.title,
      Passcode: request.passcode,
      Channel: request.channel,
    });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const itemId: number = addedItem?.data?.Id ?? addedItem?.data?.ID ?? (addedItem as any)?.Id ?? 0;

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