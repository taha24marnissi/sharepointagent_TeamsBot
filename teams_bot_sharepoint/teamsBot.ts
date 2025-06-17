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

      // Normal message flow
      const response = await axios.post("http://0.0.0.0:8000/SharepointAgent", {
        user_input: context.activity.text,
        thread_id: "1"
      }, {
        headers: { "Content-Type": "application/json" },
      });

      const data = response.data as { response?: string; status?: string; thread_id?: string };

      if (data.status === "interrupted" && data.thread_id) {
        // Send adaptive card with approve/reject
        var cardActions = [];
        cardActions.push({
            type: ActionTypes.MessageBack,
            title: 'Approve',
            value: data,
            text: 'ApproveButton'
        });
        cardActions.push({
            type: ActionTypes.MessageBack,
            title: 'Reject',
            value: data,
            text: 'RejectButton'
        });
 const card = CardFactory.heroCard(
            'Approval Card',
            data.response ?? "No response from the agent.",
            null,
            cardActions
        );
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
