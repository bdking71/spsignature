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
export declare function createPendingVerification(context: WebPartContext, request: IVerificationRequest): Promise<IVerificationResult>;
