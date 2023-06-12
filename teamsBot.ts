import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  AdaptiveCardInvokeValue,
  AdaptiveCardInvokeResponse,
  TaskModuleRequest,
  TaskModuleResponse,
  MessageFactory,
} from "botbuilder";
import rawWelcomeCard from "./adaptiveCards/welcome.json";
import rawLearnCard from "./adaptiveCards/learn.json";
import { AdaptiveCards } from "@microsoft/adaptivecards-tools";

export interface DataInterface {
  likeCount: number;
}

const taskModuleCard = {
  $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
  version: '1.0',
  type: 'AdaptiveCard',
  body: [
      {
          type: 'TextBlock',
          text: 'Task Module Invocation from Adaptive Card',
          weight: 'bolder',
          size: 3
      }
  ],
  actions:
  {
    type: 'Action.Submit',
    title: "Trigger Task Module",
    data: { msteams: { type: 'task/fetch' }, data: 500 }
  }
  // })
  // "task": {
  //   "type": "continue",
  //   "value": {
  //     "title": "Task module title",
  //     "height": 500,
  //     "width": "medium",
  //     "url": "https://contoso.com/msteams/taskmodules/newcustomer",
  //     "fallbackUrl": "https://contoso.com/msteams/taskmodules/newcustomer"
  //   }
  // }
};

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

      // Trigger command by IM text
      switch (txt) {
        case "welcome": {
          const card = AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        case "learn": {
          this.likeCountObj.likeCount = 0;
          const card = AdaptiveCards.declare<DataInterface>(rawLearnCard).render(this.likeCountObj);
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        case "frame": {
          const heroCard = CardFactory.heroCard(
            'Task Module Invocation from Hero Card',
            '',
            null, // No images
          [{
              type: 'invoke',
              title: "Show Task Module",
              value: {
                  type: 'task/fetch',
                  data: 500
              }
            }]
          );

          const reply = MessageFactory.list([heroCard]);
          await context.sendActivity(reply);

          // const reply = MessageFactory.list([taskModuleCard as any]);
          // await context.sendActivity(reply);
          break;
        }
        /**
         * case "yourCommand": {
         *   await context.sendActivity(`Add your response here!`);
         *   break;
         * }
         */
      }
  
      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
      }
      await next();
    });
  }

  override handleTeamsTaskModuleFetch(context: TurnContext, taskModuleRequest: TaskModuleRequest): Promise<TaskModuleResponse> {
    // Called when the user selects an options from the displayed HeroCard or
    // AdaptiveCard.  The result is the action to perform.

    console.log(`WE ARE HEREEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEERE`);

    const cardTaskFetchValue = taskModuleRequest.data.data;
    var taskInfo = {
      url: "https://www.bing.com/",
      fallbackUrl: "https://www.bing.com/",
      height: 510,
      width: 450,
      title: "A Task Module",
    };

    return Promise.resolve({
      task: {
        type: 'continue',
        value: taskInfo,
      }
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
      this.likeCountObj.likeCount++;
      const card = AdaptiveCards.declare<DataInterface>(rawLearnCard).render(this.likeCountObj);
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(card)],
      });
      return { statusCode: 200, type: undefined, value: undefined };
    }
  }
}
