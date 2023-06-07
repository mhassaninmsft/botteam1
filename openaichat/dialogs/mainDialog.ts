// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import {
  ConfirmPrompt,
  DialogSet,
  DialogTurnStatus,
  OAuthPrompt,
  WaterfallDialog,
} from "botbuilder-dialogs";
import { LogoutDialog } from "./logoutDialog";
import { SimpleGraphClient } from "../util/simpleGraphClient";
import { CardFactory } from "botbuilder-core";
import { DialogContext } from "botbuilder-dialogs";
import { TurnContext, StatePropertyAccessor } from "botbuilder";
import "isomorphic-fetch";
const CONFIRM_PROMPT = "ConfirmPrompt";
const MAIN_DIALOG = "MainDialog";
const MAIN_WATERFALL_DIALOG = "MainWaterfallDialog";
const OAUTH_PROMPT = "OAuthPrompt";

class MainDialog extends LogoutDialog {
  constructor() {
    super(MAIN_DIALOG, process.env.connectionName!);

    this.addDialog(
      new OAuthPrompt(OAUTH_PROMPT, {
        connectionName: process.env.connectionName!,
        text: "Please Sign In",
        title: "Sign In",
        timeout: 300000,
      })
    );
    this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
    this.addDialog(
      new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
        this.promptStep.bind(this),
        this.loginStep.bind(this),
        this.ensureOAuth.bind(this),
        this.displayToken.bind(this),
      ])
    );

    this.initialDialogId = MAIN_WATERFALL_DIALOG;
  }

  async run(
    context: TurnContext,
    accessor: StatePropertyAccessor
  ): Promise<void> {
    const dialogSet = new DialogSet(accessor);
    dialogSet.add(this);
    const dialogContext = await dialogSet.createContext(context);
    const results = await dialogContext.continueDialog();
    if (results.status === DialogTurnStatus.empty) {
      await dialogContext.beginDialog(this.id);
    }
  }

  async promptStep(stepContext: DialogContext): Promise<any> {
    try {
      return await stepContext.beginDialog(OAUTH_PROMPT);
    } catch (err) {
      console.error(err);
    }
  }

  async loginStep(stepContext: DialogContext): Promise<any> {
    // stepContext.
    const tokenResponse = (stepContext as any).result; //TODO: CHECK ME
    if (!tokenResponse || !tokenResponse.token) {
      await stepContext.context.sendActivity(
        "Login was not successful please try again."
      );
    } else {
      const client = new SimpleGraphClient(tokenResponse.token);
      const me = await client.getMe();
      const title = me ? me.jobTitle : "UnKnown";
      await stepContext.context.sendActivity(
        `You're logged in as ${me.displayName} (${me.userPrincipalName}); your job title is: ${title}; your photo is: `
      );
      const photoBase64 = await client.GetPhotoAsync(tokenResponse.token);
      const card = CardFactory.thumbnailCard(
        "",
        CardFactory.images([photoBase64])
      );
      await stepContext.context.sendActivity({ attachments: [card] });
      return await stepContext.prompt(
        CONFIRM_PROMPT,
        "Would you like to view your token?"
      );
    }
    return await stepContext.endDialog();
  }

  async ensureOAuth(stepContext: DialogContext): Promise<any> {
    await stepContext.context.sendActivity("Thank you.");

    const result = (stepContext as any).result;
    if (result) {
      return await stepContext.beginDialog(OAUTH_PROMPT);
    }
    return await stepContext.endDialog();
  }

  async displayToken(stepContext: DialogContext): Promise<any> {
    const tokenResponse = (stepContext as any).result;
    if (tokenResponse && tokenResponse.token) {
      await stepContext.context.sendActivity(
        `Here is your token ${tokenResponse.token}`
      );
    }
    return await stepContext.endDialog();
  }
}

export { MainDialog };
``;
