import {
  AuthConfiguration,
  authorizeJWT,
  CloudAdapter,
  loadAuthConfigFromEnv,
  Request,
  TurnContext,
} from "@microsoft/agents-hosting";
import express, { Response } from "express";

import { teamsBot } from "./teamsBot";

// Create authentication configuration
const authConfig: AuthConfiguration = loadAuthConfigFromEnv();

// Create adapter
const adapter = new CloudAdapter(authConfig);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Only send error message for user messages, not for other message types so the bot doesn't spam a channel or chat.
  if (context.activity.type === "message") {
    // Send a trace activity
    await context.sendTraceActivity(
      "OnTurnError Trace",
      `${error}`,
      "https://www.botframework.com/schemas/error",
      "TurnError"
    );

    // Send a message to the user
    await context.sendActivity(`The bot encountered unhandled error:\n ${error.message}`);
    await context.sendActivity("To continue to run this bot, please fix the bot source code.");
  }
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create express application
const server = express();
server.use(express.json());
server.use(authorizeJWT(authConfig));

// Add multer for file upload handling
import multer from 'multer';
import path from 'path';

// Configure multer for memory storage
const upload = multer({ 
  storage: multer.memoryStorage(),
  limits: { 
    fileSize: 10 * 1024 * 1024 // 10MB limit
  }
});

// Listen for incoming requests.
server.post("/api/messages", async (req: Request, res: Response) => {
  await adapter.process(req, res, async (context) => {
    await teamsBot.run(context);
  });
});

// Add dedicated file upload endpoint as fallback
server.post("/api/upload", upload.single('file'), async (req: any, res: any) => {
  try {
    console.log(`📤 Direct file upload request received`);
    
    if (!req.file) {
      res.status(400).json({ error: "No file provided" });
      return;
    }

    const { originalname, buffer, mimetype } = req.file;
    const destination = req.body.destination || req.query.destination || "Upload to Documents library in SharePoint";
    
    console.log(`📄 File: ${originalname}, Size: ${buffer.length}, Type: ${mimetype}`);
    
    // Forward to backend API
    const backendUrl = process.env.API_URL || "http://localhost:8000";
    const FormData = require('form-data');
    const axios = require('axios');
    
    const formData = new FormData();
    formData.append('file', buffer, { 
      filename: originalname, 
      contentType: mimetype 
    });

    const uploadResponse = await axios.post(
      `${backendUrl}/upload`,
      formData,
      {
        headers: {
          ...formData.getHeaders(),
        },
        params: {
          destination: destination
        },
        timeout: 30000
      }
    );

    res.json(uploadResponse.data);
    
  } catch (error: any) {
    console.error("Direct upload error:", error);
    res.status(500).json({ 
      error: "Upload failed", 
      details: error.message 
    });
  }
});

// Add health check endpoint
server.get("/health", (req, res) => {
  res.json({ 
    status: "healthy", 
    timestamp: new Date().toISOString(),
    backendUrl: process.env.API_URL || "http://localhost:8000",
    botId: authConfig.clientId 
  });
});

// Start the server
const port = process.env.PORT || 3978;
const backendUrl = process.env.API_URL || "http://localhost:8000";

server
  .listen(port, () => {
    console.log(`🤖 SharePoint AI Assistant Teams Bot Started!`);
    console.log(`📡 Server: http://localhost:${port}`);
    console.log(`🔗 Backend API: ${backendUrl}`);
    console.log(`🆔 Bot ID: ${authConfig.clientId}`);
    console.log(`🐛 Debug Mode: ${process.env.DEBUG || 'false'}`);
    console.log(`✅ Ready to receive messages!`);
  })
  .on("error", (err) => {
    console.error("❌ Server startup error:", err);
    process.exit(1);
  });
