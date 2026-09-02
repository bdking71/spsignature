import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IEnsurePendingVerificationsListProps {
    context: WebPartContext;
}
export declare const ensurePendingVerificationsList: (props: IEnsurePendingVerificationsListProps) => Promise<void>;
