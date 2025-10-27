import { ActivityTypes } from "@microsoft/agents-activity";
import {
  AgentApplication,
  MemoryStorage,
  TurnContext,
  TurnState,
} from "@microsoft/agents-hosting"; 
import axios from "axios";
import FormData from "form-data";
import { CardFactory } from "botbuilder";
import { Activity } from "@microsoft/agents-activity";
import { createApprovalCard } from "./cards/approvalCard";
import { createLoadingCard, createCompletionCard } from "./cards/loadingCard";
import { handleApprovalDecision } from "./helpers/approvalHelpers";

interface ConversationState {
  count: number;
} 

type ApplicationTurnState = TurnState<ConversationState>;

// Backend API interfaces
interface ChatRequest {
  message: string;
  thread_id: string;
}

interface ChatResponse {
  response: string;
  status?: string;
}

interface UploadResponse {
  response: string;
}

// Define storage and application
const storage = new MemoryStorage();
export const teamsBot = new AgentApplication<ApplicationTurnState>({
  storage,
});

// Helper: resolve backend base URL with validation
function getApiUrl(): string {
  const apiUrl = process.env.API_URL || "http://localhost:8000";
  
  // Validate API URL format
  try {
    new URL(apiUrl);
    return apiUrl;
  } catch (error) {
    console.error(`Invalid API_URL: ${apiUrl}. Using default.`);
    return "http://localhost:8000";
  }
}

// Helper: Check if backend is available
async function checkBackendHealth(): Promise<boolean> {
  try {
    const apiUrl = getApiUrl();
    const response = await axios.get(`${apiUrl}/`, { timeout: 5000 });
    return response.status === 200;
  } catch (error) {
    console.error("Backend health check failed:", error);
    return false;
  }
}

teamsBot.conversationUpdate(
  "membersAdded",
  async (context: TurnContext, state: ApplicationTurnState) => {
    // Check backend availability on startup
    const backendAvailable = await checkBackendHealth();
    
    if (!backendAvailable) {
      await context.sendActivity(
        `⚠️ **AI Assistant Starting Up**\n\n` +
        `I'm connecting to my AI backend services. Please wait a moment and try again.\n\n` +
        `If this persists, please contact your administrator.`
      );
      return;
    }

    await context.sendActivity(
      `🤖 **Hi there! I'm your AI Assistant with SharePoint Integration!**\n\n` +
      `I can help you with:\n\n` +
      `💬 **General Chat (like ChatGPT)**:\n` +
      `• "Hello, how are you?"\n` +
      `• "Explain quantum physics"\n` +
      `• "Help me write a Python function"\n` +
      `• "What are some good books to read?"\n\n` +
      `📊 **SharePoint Operations**:\n` +
      `• "Get info for site contoso.sharepoint.com/sites/marketing"\n` +
      `• "List document libraries in our SharePoint site"\n` +
      `• "Show me documents in the shared library"\n` +
      `• "Create item in list: abc123 Title: New Task Status: Active"\n\n` +
      `📄 **File Upload to SharePoint**:\n` +
      `Just attach a file and tell me where to put it!\n` +
      `Example: "Upload to site marketing in Documents"\n\n` +
      `🔧 **Commands**:\n` +
      `• Type "/reset" to clear conversation history\n\n` +
      `**Ready to chat or help with SharePoint! What can I do for you?**`
    );
  }
);

// Helper function to handle file attachments
async function handleFileAttachment(
  context: TurnContext,
  attachment: any,
  destination?: string
): Promise<string> {
  try {
    // Enhanced attachment processing with better logging
    console.log(`📥 Processing attachment: ${attachment.name || 'Unnamed file'}`);
    console.log(`📎 Content Type: ${attachment.contentType || 'Unknown'}`);
    console.log(`📏 Content URL: ${attachment.contentUrl ? 'Available' : 'Not available'}`);
    console.log(`📦 Direct Content: ${attachment.content ? 'Available' : 'Not available'}`);
    
    let fileContent: Buffer;
    
    try {
      // Try to get content from the attachment with improved handling
      if (attachment.content) {
        console.log(`📋 Using direct content attachment`);
        // Direct content available
        if (attachment.content instanceof Buffer) {
          fileContent = attachment.content;
        } else if (typeof attachment.content === 'string') {
          // Try to detect if it's base64 encoded
          try {
            fileContent = Buffer.from(attachment.content, 'base64');
          } catch (e) {
            // If base64 fails, treat as plain text
            fileContent = Buffer.from(attachment.content, 'utf8');
          }
        } else if (attachment.content instanceof ArrayBuffer) {
          fileContent = Buffer.from(attachment.content);
        } else {
          // Fallback: stringify and convert
          fileContent = Buffer.from(JSON.stringify(attachment.content));
        }
      } else if (attachment.contentUrl) {
        console.log(`🌐 Downloading from content URL: ${attachment.contentUrl}`);
        // Download from URL with enhanced error handling
        const downloadResponse = await axios.get(attachment.contentUrl, { 
          responseType: 'arraybuffer',
          timeout: 30000,
          maxRedirects: 3,
          validateStatus: (status) => status < 400
        });
        fileContent = Buffer.from(downloadResponse.data);
        console.log(`✅ Downloaded ${fileContent.length} bytes`);
      } else {
        console.error(`❌ No content source available for attachment`);
        return "❌ Cannot access attachment content - no content URL or direct content available";
      }
    } catch (downloadError: any) {
      console.error(`❌ Failed to get attachment content:`, downloadError);
      return `❌ Failed to download attachment: ${downloadError.message}`;
    }

    // Validate file size (10MB limit - same as backend)
    if (fileContent.length > 10 * 1024 * 1024) {
      return "❌ File too large (maximum 10MB allowed)";
    }

    // Get file info
    const fileName = attachment?.name || "upload";
    const contentType = attachment?.contentType || "application/octet-stream";
    
    // Validate file extension (same as backend)
    const allowedExtensions = ['.pdf', '.docx', '.xlsx', '.pptx', '.txt', '.jpg', '.png', '.gif', '.zip'];
    const fileExtension = fileName.toLowerCase().substring(fileName.lastIndexOf('.'));
    
    if (!allowedExtensions.includes(fileExtension)) {
      return `❌ File type not supported. Allowed types: ${allowedExtensions.join(', ')}`;
    }

    // Prepare multipart form data for backend API
    const formData = new FormData();
    formData.append("file", fileContent, { 
      filename: fileName, 
      contentType: contentType 
    });

    // Use destination from message or intelligent default
    const uploadDestination = destination || 
      "Upload to Documents library in SharePoint";

    // Upload to backend using the /upload endpoint
    const apiUrl = getApiUrl();
    console.log(`📤 Uploading ${fileName} to backend: ${apiUrl}/upload`);
    
    const uploadResponse = await axios.post(
      `${apiUrl}/upload`,
      formData,
      { 
        headers: {
          ...formData.getHeaders(),
        },
        params: {
          destination: uploadDestination
        },
        timeout: 30000 // 30 second timeout for uploads
      }
    );

    const result = uploadResponse.data as UploadResponse;
    return result.response || "✅ File uploaded successfully!";
    
  } catch (error: any) {
    console.error("File upload error:", error);
    
    // Provide helpful error messages based on error type
    if (error.code === 'ECONNREFUSED') {
      return "❌ Cannot connect to AI backend. Please try again later.";
    } else if (error.response?.status === 400) {
      return `❌ Upload error: ${error.response?.data?.detail || "Invalid file or destination"}`;
    } else if (error.response?.status === 503) {
      return "🔄 AI backend is starting up. Please try again in a moment.";
    } else if (error.response?.data?.detail) {
      return `❌ ${error.response.data.detail}`;
    } else {
      return `❌ Upload failed: ${error.message}`;
    }
  }
}

// Listen for approval button clicks (Action.Submit from Adaptive Cards)
teamsBot.activity(
  "invoke" as any,
  async (context: TurnContext, state: ApplicationTurnState) => {
    try {
      // Handle adaptive card action submits
      if (context.activity.name === "adaptiveCard/action") {
        const actionData = (context.activity.value as any)?.action;
        const conversationId = context.activity.conversation.id;
        
        console.log(`🔘 Adaptive card action received: ${actionData}`);
        
        if (actionData === "approval_proceed" || actionData === "approval_cancel") {
          const decision = actionData === "approval_proceed" ? "Proceed" : "Cancel";
          
          // Update the card to show it's disabled
          const disabledCard = createApprovalCard(
            "Processing your decision...",
            true,
            decision
          );
          
          // Update the original card message
          try {
            const cardMessage = {
              type: 'message',
              attachments: [CardFactory.adaptiveCard(disabledCard)]
            };
            
            // Update using the activity ID
            if (context.activity.replyToId) {
              await context.updateActivity({
                ...cardMessage,
                id: context.activity.replyToId
              } as any);
            }
          } catch (updateError) {
            console.log("Could not update card (may not be supported in this channel):", updateError);
          }
          
          // Process the approval decision
          await context.sendActivity(`⏳ ${decision === "Proceed" ? "Processing your approval" : "Cancelling operation"}...`);
          const apiUrl = getApiUrl();
          const result = await handleApprovalDecision(apiUrl, conversationId, decision);
          await context.sendActivity(result);
          return;
        }
      }
    } catch (error: any) {
      console.error("Invoke handler error:", error);
      await context.sendActivity(`❌ Error processing action: ${error.message}`);
    }
  }
);

// Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
teamsBot.activity(
  ActivityTypes.Message,
  async (context: TurnContext, state: ApplicationTurnState) => {
    try {
      const conversationId = context.activity.conversation.id;
      const apiUrl = getApiUrl();
      const messageText = context.activity.text || "";
      
      // Check if this is a response to an approval card (value property contains action data)
      if (context.activity.value && (context.activity.value as any).action) {
        const action = (context.activity.value as any).action;
        console.log(`🔘 Approval action detected: ${action}`);
        
        if (action === "approval_proceed" || action === "approval_cancel") {
          const decision = action === "approval_proceed" ? "Proceed" : "Cancel";
          
          // Update the card to show it's disabled
          const disabledCard = createApprovalCard(
            "Processing your decision...",
            true,
            decision
          );
          
          // Try to update the original card
          try {
            const cardMessage = {
              type: 'message',
              attachments: [CardFactory.adaptiveCard(disabledCard)]
            };
            
            if (context.activity.replyToId) {
              await context.updateActivity({
                ...cardMessage,
                id: context.activity.replyToId
              } as any);
            }
          } catch (updateError) {
            console.log("Could not update card (may not be supported in this channel):", updateError);
          }
          
          await context.sendActivity(`⏳ ${decision === "Proceed" ? "Processing your approval" : "Cancelling operation"}...`);
          const result = await handleApprovalDecision(apiUrl, conversationId, decision);
          await context.sendActivity(result);
          return;
        }
      }

      // Enhanced file attachment detection with detailed logging
      console.log(`🔍 Checking for attachments...`);
      console.log(`📎 Total attachments: ${context.activity.attachments?.length || 0}`);
      
      if (context.activity.attachments && context.activity.attachments.length > 0) {
        context.activity.attachments.forEach((att, index) => {
          console.log(`   📄 Attachment ${index}:`);
          console.log(`      - Name: ${att.name || 'No name'}`);
          console.log(`      - ContentType: ${att.contentType || 'No content type'}`);
          console.log(`      - ContentUrl: ${att.contentUrl ? 'Available' : 'Not available'}`);
          console.log(`      - Content: ${att.content ? 'Available' : 'Not available'}`);
          console.log(`      - Size: ${att.content ? (typeof att.content === 'string' ? att.content.length : 'Buffer') : 'Unknown'}`);
        });
      }

      // Improved file attachment detection
      const hasRealAttachments = context.activity.attachments && 
                                context.activity.attachments.length > 0 &&
                                context.activity.attachments.some(att => {
                                  // More flexible detection logic
                                  const hasName = att.name && att.name.trim() !== '';
                                  const hasContentUrl = att.contentUrl && att.contentUrl.trim() !== '';
                                  const hasContent = att.content;
                                  const hasValidContentType = att.contentType && 
                                    !att.contentType.includes('messageCard') &&
                                    !att.contentType.includes('adaptive-card') &&
                                    !att.contentType.includes('application/vnd.microsoft.card') &&
                                    // Allow Teams file download info - this is a valid file attachment
                                    !(att.contentType.includes('application/vnd.microsoft.teams') && 
                                      !att.contentType.includes('file.download.info'));
                                  
                                  const isValidAttachment = (hasName || hasContentUrl) && 
                                                          hasValidContentType && 
                                                          (hasContentUrl || hasContent);
                                  
                                  console.log(`      ✅ Attachment validation: ${isValidAttachment ? 'VALID' : 'INVALID'}`);
                                  console.log(`         - Has name: ${hasName}`);
                                  console.log(`         - Has content URL: ${hasContentUrl}`);
                                  console.log(`         - Has content: ${hasContent ? 'Yes' : 'No'}`);
                                  console.log(`         - Valid content type: ${hasValidContentType}`);
                                  
                                  return isValidAttachment;
                                });

      if (hasRealAttachments) {
        console.log(`📁 Processing file upload request`);
        
        const uploadResults: string[] = [];
        const destinationText = messageText.trim() || "Upload to Documents library in SharePoint";

        for (const attachment of context.activity.attachments) {
          // Enhanced filtering logic for system attachments
          const isSystemAttachment = !attachment.name || 
              attachment.contentType?.includes('messageCard') || 
              attachment.contentType?.includes('adaptive-card') ||
              attachment.contentType?.includes('application/vnd.microsoft.card') ||
              // Allow Teams file download info - this is a valid file attachment
              (attachment.contentType?.includes('application/vnd.microsoft.teams') && 
               !attachment.contentType?.includes('file.download.info')) ||
              (!attachment.contentUrl && !attachment.content);
          
          if (isSystemAttachment) {
            console.log(`⏭️ Skipping system attachment: ${attachment.contentType}`);
            continue;
          }

          console.log(`📤 Processing file attachment: ${attachment.name}`);
          const result = await handleFileAttachment(context, attachment, destinationText);
          uploadResults.push(result);
        }

        if (uploadResults.length > 0) {
          await context.sendActivity(uploadResults.join("\n\n"));
          return;
        }
      }

      // Check for upload intent in message text (fallback for missed attachments)
      const uploadKeywords = ['upload', 'file', 'document', 'attach', 'share'];
      const hasUploadIntent = uploadKeywords.some(keyword => 
        messageText.toLowerCase().includes(keyword)
      );
      
      if (hasUploadIntent && (!messageText.trim() || messageText.trim().length < 10)) {
        await context.sendActivity(
          `🤔 It looks like you want to upload a file, but I don't see any attachments.\n\n` +
          `**To upload files:**\n` +
          `1. Drag and drop your file into this chat\n` +
          `2. Or click the attachment button (📎) to select a file\n` +
          `3. Add a message like "Upload to SharePoint site: [your-site-url]"\n\n` +
          `**Supported file types:** PDF, DOCX, XLSX, PPTX, TXT, JPG, PNG, GIF, ZIP (max 10MB)`
        );
        return;
      }

      // Skip empty messages
      if (!messageText.trim()) {
        await context.sendActivity("Please send me a message or attach a file to upload to SharePoint!");
        return;
      }

      console.log(`💬 Processing chat message from conversation ${conversationId}: ${messageText}`);
      console.log(`📎 Attachments found: ${context.activity.attachments?.length || 0}`);
      if (context.activity.attachments?.length > 0) {
        context.activity.attachments.forEach((att, index) => {
          console.log(`   Attachment ${index}: ${att.contentType} - ${att.name}`);
        });
      }

      // Prepare request for your backend /chat endpoint
      const chatRequest: ChatRequest = {
        message: messageText,
        thread_id: conversationId
      };

      // Show typing indicator immediately
      await context.sendActivity({ type: ActivityTypes.Typing } as any);

      // Track request start time
      const startTime = Date.now();
      let progressMessageSent = false;
      let loadingCardActivityId: string | undefined;
      let progressInterval: NodeJS.Timeout | null = null;

      // Set up progress indicator for long-running requests
      progressInterval = setInterval(async () => {
        const elapsed = Math.floor((Date.now() - startTime) / 1000);
        
        if (elapsed >= 20 && !progressMessageSent) {
          // After 20 seconds, send a beautiful loading card with ProgressBar
          progressMessageSent = true;
          
          const loadingCard = createLoadingCard(
            "This request is taking longer than usual. Processing complex operations ...",
            
          );
          
          const cardAttachment = CardFactory.adaptiveCard(loadingCard);
          const loadingMessage = await context.sendActivity({
            type: 'message',
            attachments: [cardAttachment]
          } as any);
          
          // Log the raw response to assist with adapter compatibility/debugging
          console.log('📝 loading card sendActivity result:', loadingMessage);

          // Store the activity ID for potential cleanup (adapter may return id or activity object)
          loadingCardActivityId = loadingMessage?.id || (loadingMessage as any)?.activityId || undefined;
        }
      }, 5000); // Check every 5 seconds

      try {
        // Send to your AI backend using the /chat endpoint
        const response = await axios.post(
          `${apiUrl}/chat`,
          chatRequest,
          {
            headers: { 
              "Content-Type": "application/json" 
            },
            timeout: 0 // No timeout for long-running requests
          }
        );

        const chatResponse = response.data as ChatResponse;

        // Clear the progress interval once we get a response
        if (progressInterval) {
          clearInterval(progressInterval);
          progressInterval = null;
        }

        // Check if approval is required (interrupted status)
        if (chatResponse.status === "interrupted") {
          console.log(`⚠️ Approval required - showing approval card`);
          
          // Create and send approval card
          const approvalCard = createApprovalCard(chatResponse.response);
          const cardAttachment = CardFactory.adaptiveCard(approvalCard);
          
          await context.sendActivity({
            type: 'message',
            attachments: [cardAttachment]
          } as any);
          
          return; // Exit here, wait for user's approval decision
        }

        // Send the AI response back to Teams for normal flow
        const responseText = chatResponse.response || "I'm here to help you with anything you need!";
        
        console.log(`🤖 AI Response length: ${responseText.length} characters`);
        
        // Log completion time
        const totalTime = Math.floor((Date.now() - startTime) / 1000);
        console.log(`✅ Request completed in ${totalTime} seconds`);
        
        // If loading card was shown, UPDATE it to green completion state
        if (progressMessageSent && loadingCardActivityId) {
          try {
            const completionCard = createCompletionCard(totalTime);
            const completionAttachment = CardFactory.adaptiveCard(completionCard);

            // Create an Activity instance so applyConversationReference exists and adapter normalization works
            // Build a complete-ish activity object for updateActivity. Activity.fromObject validates
            // the shape (Zod), so ensure `type` is present and include the conversation id.
            const updatedActivityPayload: any = {
              id: loadingCardActivityId,
              type: 'message',
              attachments: [completionAttachment],
              conversation: { id: context.activity.conversation.id }
            };

            const updatedActivity = Activity.fromObject(updatedActivityPayload);
            await context.updateActivity(updatedActivity as any);
            console.log('✅ Updated loading card in-place to completion state');
          } catch (updateErr) {
            console.log('In-place update failed, falling back to send-only:', updateErr);
            try {
              const completionCard = createCompletionCard(totalTime);
              const completionAttachment = CardFactory.adaptiveCard(completionCard);
              await context.sendActivity({ type: 'message', attachments: [completionAttachment] } as any);
            } catch (sendErr) {
              console.log('Failed to send completion card as fallback:', sendErr);
            }
          }
        } else if (progressMessageSent) {
          // Progress shown but we don't have an id - just send completion card
          try {
            const completionCard = createCompletionCard(totalTime);
            const completionAttachment = CardFactory.adaptiveCard(completionCard);
            await context.sendActivity({ type: 'message', attachments: [completionAttachment] } as any);
          } catch (sendErr) {
            console.log('Failed to send completion card:', sendErr);
          }
        }
        
        await context.sendActivity(responseText);
      } catch (requestError: any) {
        // Clear the progress interval on error
        if (progressInterval) {
          clearInterval(progressInterval);
          progressInterval = null;
        }
        
        // Re-throw to be caught by outer error handler
        throw requestError;
      }
      
    } catch (error: any) {
      console.error("Bot message processing error:", error);

      // Provide specific error messages based on error type
      let errorMessage = "❌ Sorry, I encountered an error processing your request.";

      if (error.code === 'ECONNREFUSED') {
        errorMessage = "❌ Cannot connect to AI backend. The service may be starting up. Please try again in a moment.";
      } else if (error.code === 'ETIMEDOUT') {
        errorMessage = "⏱️ Request timed out. Your message may be too complex. Please try a shorter question.";
      } else if (error.response?.status === 503) {
        errorMessage = "🔄 The AI assistant is starting up. Please wait a moment and try again.";
      } else if (error.response?.status === 500) {
        errorMessage = "⚠️ There was a server error. Please try again later or contact your administrator.";
      } else if (error.response?.status === 400) {
        errorMessage = `❌ Request error: ${error.response?.data?.detail || "Invalid request format"}`;
      } else if (error.response?.data?.detail) {
        errorMessage = `❌ ${error.response.data.detail}`;
      } else if (error.message) {
        errorMessage = `❌ Error: ${error.message}`;
      }

      await context.sendActivity(errorMessage);
    }
  }
);
