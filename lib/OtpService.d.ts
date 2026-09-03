/**
 * @file OtpService.ts
 * Provides methods for dispatching One-Time Passcodes (OTP) for
 * Two-Factor Authentication via SharePoint and external channels.
 */
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
/**
 * Queues an OTP dispatch request in SharePoint to notify the user via Email/Teams.
 */
export declare function dispatchOtpPasscode(context: WebPartContext, request: IOtpDispatchRequest, listTitle?: string): Promise<IOtpDispatchResult>;
