import { ActivityTypes } from "@microsoft/agents-activity";
import {
  AgentApplication,
  MemoryStorage,
  TurnContext,
  TurnState,
} from "@microsoft/agents-hosting"; 
import axios from "axios";
import FormData from "form-data";

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

// Listen for user to say '/reset' and then delete conversation state
teamsBot.message("/reset", async (context: TurnContext, state: ApplicationTurnState) => {
  state.deleteConversationState();
  await context.sendActivity("Ok I've deleted the current conversation state.");
});

// Debug command to check attachments
teamsBot.message("/debug", async (context: TurnContext, state: ApplicationTurnState) => {
  const attachmentInfo = {
    count: context.activity.attachments?.length || 0,
    attachments: context.activity.attachments?.map(att => ({
      name: att.name,
      contentType: att.contentType,
      hasContentUrl: !!att.contentUrl,
      hasContent: !!att.content
    })) || []
  };
  
  await context.sendActivity(`**Debug Info:**\n\`\`\`json\n${JSON.stringify(attachmentInfo, null, 2)}\n\`\`\``);
}); 
 
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
    // For Microsoft 365 Agents Toolkit, we can access the attachment content directly
    // or download it using the bot framework connector
    console.log(`📥 Processing attachment: ${attachment.name}`);
    console.log(`📎 Content Type: ${attachment.contentType}`);
    console.log(`📏 Content URL: ${attachment.contentUrl ? 'Available' : 'Not available'}`);
    
    let fileContent: Buffer;
    
    // Try to get content from the attachment
    if (attachment.content) {
      // Direct content available
      if (attachment.content instanceof Buffer) {
        fileContent = attachment.content;
      } else if (typeof attachment.content === 'string') {
        fileContent = Buffer.from(attachment.content, 'base64');
      } else {
        fileContent = Buffer.from(attachment.content);
      }
    } else if (attachment.contentUrl) {
      // Download from URL
      const downloadResponse = await axios.get(attachment.contentUrl, { 
        responseType: 'arraybuffer',
        timeout: 30000 
      });
      fileContent = Buffer.from(downloadResponse.data);
    } else {
      return "❌ Cannot access attachment content";
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

// Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
teamsBot.activity(
  ActivityTypes.Message,
  async (context: TurnContext, state: ApplicationTurnState) => {
    try {
      const conversationId = context.activity.conversation.id;
      const apiUrl = getApiUrl();
      const messageText = context.activity.text || "";

      // Handle file attachments ONLY if they are actual user-uploaded files
      const hasRealAttachments = context.activity.attachments && 
                                context.activity.attachments.length > 0 &&
                                context.activity.attachments.some(att => 
                                  att.name && 
                                  att.contentType && 
                                  !att.contentType.includes('messageCard') &&
                                  !att.contentType.includes('application/vnd.microsoft') &&
                                  (att.contentUrl || att.content)
                                );

      if (hasRealAttachments) {
        console.log(`📁 Processing file upload request`);
        
        const uploadResults: string[] = [];
        const destinationText = messageText.trim() || "Upload to Documents library in SharePoint";

        for (const attachment of context.activity.attachments) {
          // Skip system attachments
          if (!attachment.name || 
              attachment.contentType?.includes('messageCard') || 
              attachment.contentType?.includes('application/vnd.microsoft')) {
            continue;
          }

          const result = await handleFileAttachment(context, attachment, destinationText);
          uploadResults.push(result);
        }

        if (uploadResults.length > 0) {
          await context.sendActivity(uploadResults.join("\n\n"));
          return;
        }
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

      // Send to your AI backend using the /chat endpoint
      const response = await axios.post(
        `${apiUrl}/chat`,
        chatRequest,
        {
          headers: { 
            "Content-Type": "application/json" 
          },
          timeout: 45000 // 45 second timeout for AI responses
        }
      );

      const chatResponse = response.data as ChatResponse;

      // Send the AI response back to Teams
      const responseText = chatResponse.response || "I'm here to help you with anything you need!";
      
      console.log(`🤖 AI Response length: ${responseText.length} characters`);
      await context.sendActivity(responseText);
      
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
