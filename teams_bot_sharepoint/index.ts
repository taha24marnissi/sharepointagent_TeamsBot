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

// Listen for incoming requests.
server.post("/api/messages", async (req: Request, res: Response) => {
  await adapter.process(req, res, async (context) => {
    await teamsBot.run(context);
  });
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
