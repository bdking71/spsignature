import { SPFI, spfi, SPFx } from "@pnp/sp";
import { IList } from "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IEnsurePendingVerificationsListProps {
  context: WebPartContext;
}

export const ensurePendingVerificationsList = async (
  props: IEnsurePendingVerificationsListProps
): Promise<void> => {
  if (!props.context) {
    throw new Error("Execution restricted: Valid WebPartContext is required.");
  }

  const sp: SPFI = spfi().using(SPFx(props.context));
  const LIST_NAME = "PendingVerifications";
  const LIST_DESCRIPTION = "Stores temporary PIN verification codes for document signatures.";
  const list: IList = sp.web.lists.getByTitle(LIST_NAME);

  let listCreated = false;

  try {
    await list();
  } catch {
    await sp.web.lists.add(LIST_NAME, LIST_DESCRIPTION, 100, true);
    listCreated = true;
  }

  if (!listCreated) {
    return;
  }

  const ensureField = async (
    fieldName: string,
    createFieldFn: () => Promise<void>
  ): Promise<void> => {
    try {
      await createFieldFn();
    } catch (creationErr: unknown) {
      throw creationErr;
    }
  };

  await ensureField("Passcode", async () => {
    await list.fields.addText("Passcode", {
      Required: true,
      Description: "Temporary 6-digit signature verification code."
    });
  });

  await ensureField("Channel", async () => {
    await list.fields.addChoice("Channel", {
      Choices: ["email", "teams", "both"],
      Required: true,
      FillInChoice: false,
      Description: "Target notification delivery route."
    });

    const channelField = list.fields.getByInternalNameOrTitle("Channel");
    await channelField.update({ DefaultValue: "teams" });
  });

  try {
    await list.update({
      ReadSecurity: 2,
      WriteSecurity: 2
    } as Record<string, unknown>);
  } catch (secErr: unknown) {
    throw secErr;
  }
};