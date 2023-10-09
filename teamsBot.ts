import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  AdaptiveCardInvokeValue,
  AdaptiveCardInvokeResponse,
} from "botbuilder";
import welcomeCard from "./adaptiveCards/welcome.json";
import feedbackCard from "./adaptiveCards/feedback.json";
import dislikeCard from "./adaptiveCards/userDislikeAction.json";
import findSupportCard from "./adaptiveCards/findSupport.json";
import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import * as teamsfxSdk from "@microsoft/teamsfx";

export interface DataInterface {
  likeCount: number;
}

export class TeamsBot extends TeamsActivityHandler {
  // record the likeCount
  likeCountObj: { likeCount: number };

  constructor() {
    super();

    this.likeCountObj = { likeCount: 0 };

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");

      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      const authProvider = new teamsfxSdk.ApiKeyProvider(
        "api-key",
        process.env.CODEGEN_API_KEY,
        teamsfxSdk.ApiKeyLocation.Header
      );
      const apiClient = teamsfxSdk.createApiClient(
        process.env.CODEGEN_API_ENDPOINT,
        authProvider
      );
      switch (txt) {
        case "welcome": {
          const card = AdaptiveCards.declareWithoutData(welcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        default: {
          const postData = { question: txt };
          const headers = { 'Content-Type': 'application/json' };
          const response = await apiClient.post('', JSON.stringify(postData), { headers: headers });
          const responseReply = response.data.reply;
          await context.sendActivity(responseReply);
          const card = AdaptiveCards.declareWithoutData(feedbackCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
        }
      }
      
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = AdaptiveCards.declareWithoutData(welcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
      }
      await next();
    });
  }

  // Invoked when an action is taken on an Adaptive Card. The Adaptive Card sends an event to the Bot and this
  // method handles that event.
  async onAdaptiveCardInvoke(
    context: TurnContext,
    invokeValue: AdaptiveCardInvokeValue
  ): Promise<AdaptiveCardInvokeResponse> {
    // The verb "userlike" is sent from the Adaptive Card defined in adaptiveCards/learn.json
    if (invokeValue.action.verb === "userlike") {
      await context.sendActivity("Thanks for the feedback! Can I help with anything else?");
    }
    else if(invokeValue.action.verb === "userDislike") {
      const card = AdaptiveCards.declareWithoutData(dislikeCard).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
    }
    else if(invokeValue.action.verb === "rephrase") {
      await context.sendActivity("Please go ahead and ask questions in the chat.");
    }
    else if(invokeValue.action.verb === "findSupport") {
      const card = AdaptiveCards.declareWithoutData(findSupportCard).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
    }
    else{
      return;
    }
    return { statusCode: 200, type: undefined, value: undefined };
  }
}
