import { ActionTypes, ActivityTypes } from "@microsoft/agents-activity";
import {
  AgentApplication,
  AttachmentDownloader,
  CardFactory,
  MessageFactory,
  MemoryStorage,
  TurnContext,
  TurnState  
} from "@microsoft/agents-hosting";
import { version } from "@microsoft/agents-hosting/package.json";
import axios from "axios";

interface ConversationState {
  count: number;
  awaitingApproval?: boolean;
}
type ApplicationTurnState = TurnState<ConversationState>;

const downloader = new AttachmentDownloader();

// Define storage and application
const storage = new MemoryStorage();
export const teamsBot = new AgentApplication<ApplicationTurnState>({
  storage,
  fileDownloaders: [downloader],
});

// Listen for user to say '/reset' and then delete conversation state
teamsBot.message("/reset", async (context: TurnContext, state: ApplicationTurnState) => {
  state.deleteConversationState();
  await context.sendActivity("Ok I've deleted the current conversation state.");
});

teamsBot.message("/count", async (context: TurnContext, state: ApplicationTurnState) => {
  const count = state.conversation.count ?? 0;
  await context.sendActivity(`The count is ${count}`);
});

teamsBot.message("/diag", async (context: TurnContext, state: ApplicationTurnState) => {
  await state.load(context, storage);
  await context.sendActivity(JSON.stringify(context.activity));
});

teamsBot.message("/state", async (context: TurnContext, state: ApplicationTurnState) => {
  await state.load(context, storage);
  await context.sendActivity(JSON.stringify(state));
});

teamsBot.message("/runtime", async (context: TurnContext, state: ApplicationTurnState) => {
  const runtime = {
    nodeversion: process.version,
    sdkversion: version,
  };
  await context.sendActivity(JSON.stringify(runtime));
});

teamsBot.conversationUpdate(
  "membersAdded",
  async (context: TurnContext, state: ApplicationTurnState) => {
    await context.sendActivity(
      `Hi there! I'm an echo bot running on Agents SDK version ${version} that will echo what you said to me.`
    );
  }
);

// Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
teamsBot.activity(
  ActivityTypes.Message,
  async (context: TurnContext, state: ApplicationTurnState) => {
    try {
      if (context.activity.text === "ApproveButton" || context.activity.text === "RejectButton") {
        // Handle the adaptive card action
        let user_input = context.activity.text;
        if (user_input === "ApproveButton") {
          user_input = "Approve";
        } else if (user_input === "RejectButton") {
          user_input = "Reject";
        }
        // Clear approval state
        state.conversation.awaitingApproval = false;
        const continueResponse = await axios.post("http://localhost:8000/SharepointAgent-continue", {
          user_input,
          thread_id: "1"
        }, {
          headers: { "Content-Type": "application/json" },
        });

        const continueData = continueResponse.data as { response?: string };
        await context.sendActivity(continueData.response ?? "No response from the agent.");
        return;
      }

      // If awaiting approval, block other messages
      if (state.conversation.awaitingApproval) {
        await context.sendActivity("You have a pending approval request. Please respond to the approval card before sending other messages.");
        return;
      }

      // Normal message flow
      const response = await axios.post("http://0.0.0.0:8000/SharepointAgent", {
        user_input: context.activity.text,
        thread_id: "1"
      }, {
        headers: { "Content-Type": "application/json" },
      });

      const data = response.data as { response?: string; status?: string; thread_id?: string };

      if (data.status === "interrupted" && data.thread_id) {
        // Set approval state
        state.conversation.awaitingApproval = true;
        // Send adaptive card with approve/reject
        // Adaptive Card JSON for approval
        const adaptiveCard = {
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.4",
          "body": [
            {
              "type": "ColumnSet",
              "columns": [
                {
                  "type": "Column",
                  "width": "auto",
                  "items": [
                    {
                      "type": "Image",
                      "url": "https://www.clipartmax.com/png/middle/118-1180913_approve-document-icons-tick-and-cross-icon.png",
                      "size": "Small",
                      "style": "Person"
                    }
                  ]
                },
                {
                  "type": "Column",
                  "width": "stretch",
                  "items": [
                    {
                      "type": "TextBlock",
                      "text": "Action Required: Approval Needed",
                      "weight": "Bolder",
                      "size": "Large"
                    },
                    {
                      "type": "TextBlock",
                      "text": "Please review the following request and choose to Approve or Reject.",
                      "isSubtle": true,
                      "wrap": true
                    }
                  ]
                }
              ]
            },
            {
              "type": "Container",
              "items": [
                {
                  "type": "TextBlock",
                  "text": data.response ?? "No response from the agent.",
                  "wrap": true,
                  "spacing": "Medium",
                  "size": "Medium"
                }
              ],
              "style": "emphasis",
              "bleed": true
            },
            {
              "type": "TextBlock",
              "text": "If you have any questions, please contact your administrator.",
              "isSubtle": true,
              "wrap": true,
              "spacing": "Medium"
            }
          ],
          "actions": [
            {
              "type": "Action.Submit",
              "title": "✅ Approve",
              "style": "positive",
              "data": {
                ...data,
                "msteams": { "type": "messageBack", "text": "ApproveButton" }
              }
            },
            {
              "type": "Action.Submit",
              "title": "❌ Reject",
              "style": "destructive",
              "data": {
                ...data,
                "msteams": { "type": "messageBack", "text": "RejectButton" }
              }
            }
          ]
        };
        const card = CardFactory.adaptiveCard(adaptiveCard);
        await context.sendActivity(MessageFactory.attachment(card));
      } else {
        await context.sendActivity(data.response ?? "No response from the agent.");
      }
    } catch (error) {
      await context.sendActivity("Agent api error.");
    }
  }
);

teamsBot.activity(/^message/, async (context: TurnContext, state: ApplicationTurnState) => {
  await context.sendActivity(`Matched with regex: ${context.activity.type}`);
});

teamsBot.activity(
  async (context: TurnContext) => Promise.resolve(context.activity.type === "message"),
  async (context, state) => {
    await context.sendActivity(`Matched function: ${context.activity.type}`);
  }
);
