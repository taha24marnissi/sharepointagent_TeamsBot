# Microsoft 365 Agents Toolkit - SharePoint AI Assistant

## Overview

This Teams bot is built specifically for **Microsoft 365 Agents Toolkit** and integrates with your FastAPI backend to provide:

- **💬 Conversational AI**: ChatGPT-like capabilities
- **📊 SharePoint Integration**: Natural language SharePoint operations  
- **📄 File Upload**: Direct document upload to SharePoint
- **🧠 Memory**: Context-aware conversations using thread IDs

## Microsoft 365 Agents Toolkit Features Used

### AgentApplication
- Modern bot framework with built-in state management
- Automatic conversation threading
- Memory storage for conversation state

### TurnContext & TurnState
- Rich context information for each user interaction
- Conversation state persistence across messages
- User and conversation metadata

### Activity Handling
- Message activities for chat interactions
- File attachment support for document uploads
- Conversation update events for welcome messages

## Development Workflow

### 1. Prerequisites
```bash
# Ensure you have the right versions
node --version  # Should be 18, 20, or 22
npm --version   # Should be 8+
```

### 2. Backend Connection
```bash
# Start your FastAPI backend first
cd /path/to/your/fastapi/backend
python main_chat.py  # Should run on http://localhost:8000
```

### 3. Configure Environment
```bash
# In your Teams bot directory
npm run setup  # Interactive setup script
# OR manually create .env with API_URL=http://localhost:8000
```

### 4. Run with M365 Agents Toolkit
```bash
# Install dependencies
npm install

# Build TypeScript
npm run build

# Debug in VS Code
# Press F5 → Select "Debug in Microsoft 365 Agents Playground"
```

## M365 Agents Toolkit Debugging

### Local Development
- **F5** in VS Code starts the M365 Agents Playground
- Auto-reloads when you make changes to TypeScript files
- Console logs appear in VS Code Debug Console

### Environment Configuration
The toolkit uses these configuration files:
- `m365agents.yml` - Main configuration
- `m365agents.local.yml` - Local development settings  
- `m365agents.playground.yml` - Playground-specific settings

### Teams App Manifest
- Located in `appPackage/manifest.json`
- Automatically processed by M365 Agents Toolkit
- Supports file uploads (`supportsFiles: true`)

## Integration Architecture

```
Microsoft Teams Client
        ↓
Microsoft 365 Agents Toolkit (Bot Framework)
        ↓
AgentApplication (Your Bot Logic)
        ↓
FastAPI Backend (AI + SharePoint)
        ↓
SharePoint / Microsoft Graph
```

## Key Differences from Standard Bot Framework

### Simplified Setup
- No need for Bot Registration in Azure manually
- M365 Agents Toolkit handles authentication
- Automatic tunneling for local development

### Enhanced Development Experience  
- Built-in playground for testing
- Hot reload during development
- Integrated debugging in VS Code

### Modern APIs
- Uses latest Bot Framework SDK features
- Simplified state management
- Better TypeScript support

## Testing Your Integration

### In M365 Agents Playground
1. **General Chat**: "Hello, how are you today?"
2. **SharePoint Query**: "List my SharePoint sites"
3. **File Upload**: Attach a file + "Upload to Documents"

### Production Deployment
The M365 Agents Toolkit provides:
- Automatic Azure resource provisioning
- CI/CD pipeline templates
- Production environment configuration

## Troubleshooting

### Common Issues
1. **Backend Not Connected**: Ensure FastAPI is running on port 8000
2. **Build Errors**: Run `npm run build` to check TypeScript
3. **Playground Won't Start**: Check VS Code M365 Agents Toolkit extension

### Debug Tips
- Use `console.log()` in your bot code - appears in Debug Console
- Check Network tab in playground for API calls
- Verify backend health at `http://localhost:8000/`

## Next Steps

1. **Production Deployment**: Use `teamsapp provision` and `teamsapp deploy`
2. **Azure Integration**: Configure production backend URLs
3. **Advanced Features**: Add more SharePoint operations
4. **Security**: Implement authentication for production use

Your bot is now ready to provide AI-powered SharePoint assistance through Microsoft Teams! 🚀