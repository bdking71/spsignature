/**
 * @file EnsurePendingVerificationsList.ts
 *
 * Utility module responsible for idempotent provisioning and configuration
 * of the SharePoint list used for Two-Factor Authentication (2FA) / OTP queueing.
 * Designed to be executed during admin installation or setup checks.
 *
 * @module EnsurePendingVerificationsList
 */

import { SPFI, spfi, SPFx } from "@pnp/sp";
import { IList } from "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export const DEFAULT_VERIFICATION_LIST_NAME = "PendingVerifications";
export const DEFAULT_VERIFICATION_LIST_DESC =
  "Stores temporary verification passcode records for digital signature workflows.";

export interface IEnsurePendingVerificationsListProps {
  context: WebPartContext;
  listTitle?: string;
  listDescription?: string;
}

/**
 * Idempotently provisions the 2FA verification list, adds required fields,
 * and enforces security. Intended to be run during admin setup/installation.
 */
export const ensurePendingVerificationsList = async (
  props: IEnsurePendingVerificationsListProps
): Promise<void> => {
  if (!props.context) {
    throw new Error("Execution restricted: Valid WebPartContext is required.");
  }

  const listTitle = props.listTitle || DEFAULT_VERIFICATION_LIST_NAME;
  const listDescription = props.listDescription || DEFAULT_VERIFICATION_LIST_DESC;

  const sp: SPFI = spfi().using(SPFx(props.context));
  let list: IList = sp.web.lists.getByTitle(listTitle);

  // 1. Ensure the List Exists (Create if missing)
  try {
    await list();
  } catch {
    // 100 = Generic List template
    await sp.web.lists.add(listTitle, listDescription, 100, true);
    list = sp.web.lists.getByTitle(listTitle);
  }

  // 2. Helper: Ensure Field Exists
  const ensureField = async (
    fieldName: string,
    createFieldFn: () => Promise<void>
  ): Promise<void> => {
    try {
      await list.fields.getByInternalNameOrTitle(fieldName)();
    } catch {
      await createFieldFn();
    }
  };

  // 3. Ensure 'Passcode' Field
  await ensureField("Passcode", async () => {
    await list.fields.addText("Passcode", {
      Required: true,
      Description: "Temporary 5-digit signature verification passcode.",
    });
  });

  // 4. Ensure 'Channel' Field
  await ensureField("Channel", async () => {
    await list.fields.addChoice("Channel", {
      Choices: ["email", "teams", "both"],
      Required: true,
      FillInChoice: false,
      Description: "Target notification delivery route (email, teams, or both).",
    });

    const channelField = list.fields.getByInternalNameOrTitle("Channel");
    await channelField.update({ DefaultValue: "email" });
  });

  // 5. Enforce Item-Level Security Settings (Admin-configured once)
  await list.update({
    ReadSecurity: 2,
    WriteSecurity: 2,
  } as Record<string, unknown>);
};