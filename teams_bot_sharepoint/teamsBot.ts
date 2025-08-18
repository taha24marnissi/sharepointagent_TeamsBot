import { ActivityTypes } from "@microsoft/agents-activity";
import {
  AgentApplication,
  AttachmentDownloader,
  CardFactory,
  MessageFactory,
  MemoryStorage,
  TurnContext,
  TurnState,
  
} from "@microsoft/agents-hosting";
import axios from "axios";
import { getApprovalAdaptiveCard } from "./approvalCard";
import { TeamsInfo } from "botbuilder";

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


teamsBot.conversationUpdate(
  "membersAdded",
  async (context: TurnContext, state: ApplicationTurnState) => {
    await context.sendActivity(
      `Hi there! I'm a SharePoint Agent bot running on Microsoft Teams.`
    );
  }
);

// Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
teamsBot.activity(
  ActivityTypes.Message,
  async (context: TurnContext, state: ApplicationTurnState) => {          
    try {
      var conversation_id = context.activity.conversation.id
      //var userId = context.activity.from.id;
      var CurrentUserMail ="taha@3rd5pw.onmicrosoft.com";
      var response = null;
      var user_inputs = JSON.stringify(context.activity.value);
      if (context.activity.text === "ApproveButton" || context.activity.text === "RejectButton") {
        // Handle the adaptive card action
        let answer_type = context.activity.text;
        if (answer_type === "ApproveButton") {
          answer_type = "Approve"; 
        } else if (answer_type === "RejectButton") {
          answer_type = "Reject";
        }
        // Clear approval state
        state.conversation.awaitingApproval = false;
        response= await axios.post(`${process.env.API_URL}/SharepointAgent-continue`, {
          user_input: user_inputs,
          thread_id: conversation_id,
          type: answer_type,
        }, {
          headers: { "Content-Type": "application/json" },
        });
      const data = response.data as { response?: string; status?: string; thread_id?: string };

        if (data.status === "interrupted" && data.thread_id) {
        // Set approval state
        state.conversation.awaitingApproval = true;
        // Use externalized Adaptive Card generator
        const locale = context.activity.locale?.split('-')[0] || 'en';        
        const adaptiveCard = getApprovalAdaptiveCard(data, locale);
        const card = CardFactory.adaptiveCard(adaptiveCard);
        await context.sendActivity(MessageFactory.attachment(card));
      } else {
        await context.sendActivity(data.response ?? "No response from the agent.");
      }
        //await context.sendActivity(continueData.response ?? "No response from the agent.");
        return;
      }

      // If awaiting approval, block other messages
      if (state.conversation.awaitingApproval) {
        await context.sendActivity("You have a pending approval request. Please respond to the approval card before sending other messages.");
        return;
      }
      var url=process.env.API_URL
      // Normal message flow
      response = await axios.post(`${url}/SharepointAgent`, {
        user_input: context.activity.text,
        thread_id: conversation_id,
        user_email: CurrentUserMail
      }, {
        headers: { "Content-Type": "application/json" },
      });

      const data = response.data as { response?: string; status?: string; thread_id?: string };

      if (data.status === "interrupted" && data.thread_id) {
        // Set approval state
        state.conversation.awaitingApproval = true;
        // Use externalized Adaptive Card generator
        const locale = context.activity.locale?.split('-')[0] || 'en';
     
        const adaptiveCard = getApprovalAdaptiveCard(data, locale);
        const card = CardFactory.adaptiveCard(adaptiveCard);
        await context.sendActivity(MessageFactory.attachment(card));
      } else {
        await context.sendActivity(data.response ?? "No response from the agent.");
      }
    } catch (error) {
      await context.sendActivity(`Agent api error: ${error.message}`);
    }
  }
);

teamsBot.activity(/^message/, async (context: TurnContext, state: ApplicationTurnState) => {
  await context.sendActivity(`Matched with regex: ${context.activity.type}`);
});
teamsBot.activity(ActivityTypes.Invoke, async (context: TurnContext, state: ApplicationTurnState) => {
  //await context.sendActivity(`Matched with regex: ${context.activity.type}`);
});
teamsBot.activity(
  async (context: TurnContext) => Promise.resolve(context.activity.type === "message"),
  async (context, state) => {
    await context.sendActivity(`Matched function: ${context.activity.type}`);
  }
);


