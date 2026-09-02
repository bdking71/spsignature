import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DeliveryChannel } from "./VerificationService";
export interface SignerContext {
    itemID: number;
    signer: string;
    payload: Record<string, unknown>;
    spContext: WebPartContext;
    channel?: DeliveryChannel;
    requireTFA?: boolean;
}
export interface SharePointAuditRecord {
    signatureHash: string;
    signatureData: string;
    signatureTimestamp: string;
    verificationItemId: number;
}
export interface AuditEnvelopeRecord {
    payload: Record<string, unknown>;
    signer: string;
    timestamp: string;
}
export declare function promptAndGenerateSecureAudit(context: SignerContext, modalTitle?: string, warningMessage?: string): Promise<SharePointAuditRecord | undefined>;
export declare function getReportableSignature(compressedSignatureData: string): string;
export declare function verifySecureAuditRecord(payloadToVerify: Record<string, unknown>, signer: string, timestamp: string, compressedSignatureData: string, storedHash: string): Promise<boolean>;
