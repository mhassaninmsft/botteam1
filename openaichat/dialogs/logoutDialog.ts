// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { ActivityTypes } from "botbuilder";
import {
  ComponentDialog,
  DialogContext,
  DialogTurnResult,
} from "botbuilder-dialogs";
import "isomorphic-fetch";
class LogoutDialog extends ComponentDialog {
  connectionName: string;

  constructor(id: string, connectionName: string) {
    super(id);
    this.connectionName = connectionName;
  }

  async onBeginDialog(
    innerDc: DialogContext,
    options: any
  ): Promise<DialogTurnResult> {
    const result = await this.interrupt(innerDc);
    if (result) {
      return result;
    }

    return await super.onBeginDialog(innerDc, options);
  }

  async onContinueDialog(innerDc: DialogContext): Promise<DialogTurnResult> {
    const result = await this.interrupt(innerDc);
    if (result) {
      return result;
    }

    return await super.onContinueDialog(innerDc);
  }

  async interrupt(
    innerDc: DialogContext
  ): Promise<DialogTurnResult | undefined> {
    if (innerDc.context.activity.type === ActivityTypes.Message) {
      const text = innerDc.context.activity.text.toLowerCase();
      if (text === "logout") {
        // innerDc.context.adapter.

        const userTokenClient = innerDc.context.turnState.get(
          (innerDc.context.adapter as any).UserTokenClientKey
        );
        // const userTokenClient = innerDc.context.turnState.get(innerDc.context.adapter.ConnectorClientKey);

        const { activity } = innerDc.context;
        await userTokenClient.signOutUser(
          activity.from.id,
          this.connectionName,
          activity.channelId
        );

        await innerDc.context.sendActivity("You have been signed out.");
        return await innerDc.cancelAllDialogs();
      }
    }
  }
}

export { LogoutDialog };
