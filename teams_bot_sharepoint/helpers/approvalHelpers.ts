/**
 * Approval Workflow Helpers
 * 
 * Handles the approval continuation workflow with the backend
 */

import axios from "axios";

export interface ContinueRequest {
  thread_id: string;
  user_input: string;
  type: string;
}

export interface ChatResponse {
  response: string;
  status?: string;
}

/**
 * Handles approval decision and calls the /chat-continue endpoint
 * @param apiUrl - Backend API URL
 * @param conversationId - Thread ID for the conversation
 * @param decision - User's decision ("Proceed" or "Cancel")
 * @returns Response message from backend
 */
export async function handleApprovalDecision(
  apiUrl: string,
  conversationId: string,
  decision: string
): Promise<string> {
  try {
    const continueRequest: ContinueRequest = {
      thread_id: conversationId,
      user_input: decision,
      type: "approval"
    };

    console.log(`🔄 Sending approval decision: ${decision} for thread: ${conversationId}`);

    const response = await axios.post(
      `${apiUrl}/chat-continue`,
      continueRequest,
      {
        headers: { 
          "Content-Type": "application/json" 
        },
        timeout: 45000
      }
    );

    const continueResponse = response.data as ChatResponse;
    return continueResponse.response || "Operation completed.";
    
  } catch (error: any) {
    console.error("Approval continuation error:", error);
    
    if (error.response?.status === 404) {
      return "❌ No pending approval found. The operation may have expired or already been processed.";
    } else if (error.response?.data?.detail) {
      return `❌ ${error.response.data.detail}`;
    } else {
      return `❌ Failed to process approval: ${error.message}`;
    }
  }
}
