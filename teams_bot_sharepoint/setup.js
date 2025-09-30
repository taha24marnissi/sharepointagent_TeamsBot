#!/usr/bin/env node

/**
 * Setup script for SharePoint AI Assistant Teams Bot
 * This script helps configure the environment for connecting to your FastAPI backend
 */

const fs = require('fs');
const path = require('path');
const readline = require('readline');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

async function askQuestion(question) {
  return new Promise((resolve) => {
    rl.question(question, (answer) => {
      resolve(answer);
    });
  });
}

async function setup() {
  console.log('\n🤖 SharePoint AI Assistant Teams Bot Setup\n');
  console.log('This script will help you configure your bot to connect to your FastAPI backend.\n');

  // Check if .env exists
  const envPath = path.join(__dirname, '.env');
  const envExamplePath = path.join(__dirname, '.env.example');
  
  if (fs.existsSync(envPath)) {
    const overwrite = await askQuestion('❓ .env file already exists. Overwrite? (y/N): ');
    if (overwrite.toLowerCase() !== 'y') {
      console.log('✅ Setup cancelled. Existing .env file preserved.');
      rl.close();
      return;
    }
  }

  // Get backend URL
  const defaultBackend = 'http://localhost:8000';
  const backendUrl = await askQuestion(`🔗 Backend API URL (${defaultBackend}): `) || defaultBackend;

  // Test backend connection
  console.log(`\n🔄 Testing connection to ${backendUrl}...`);
  
  try {
    const axios = require('axios');
    const response = await axios.get(`${backendUrl}/`, { timeout: 5000 });
    console.log('✅ Backend connection successful!');
    console.log(`   Service: ${response.data.service}`);
    console.log(`   Version: ${response.data.version}`);
  } catch (error) {
    console.log('⚠️  Could not connect to backend. Please make sure it\'s running.');
    console.log(`   Error: ${error.message}`);
    
    const continueAnyway = await askQuestion('Continue with setup anyway? (y/N): ');
    if (continueAnyway.toLowerCase() !== 'y') {
      console.log('❌ Setup cancelled.');
      rl.close();
      return;
    }
  }

  // Create .env file
  let envContent = '';
  
  if (fs.existsSync(envExamplePath)) {
    envContent = fs.readFileSync(envExamplePath, 'utf8');
    envContent = envContent.replace('API_URL=http://localhost:8000', `API_URL=${backendUrl}`);
  } else {
    envContent = `# SharePoint AI Assistant Teams Bot Configuration
API_URL=${backendUrl}

# Teams Bot Configuration (auto-populated by Teams Toolkit)
BOT_ID=
BOT_PASSWORD=
TEAMS_APP_ID=

# Azure Configuration (auto-populated by Teams Toolkit)
AZURE_SUBSCRIPTION_ID=
AZURE_RESOURCE_GROUP_NAME=

# Debug Configuration
DEBUG=false
LOG_LEVEL=info
`;
  }

  fs.writeFileSync(envPath, envContent);
  console.log(`\n✅ Configuration saved to .env`);
  console.log(`   Backend URL: ${backendUrl}`);
  
  console.log('\n📋 Next Steps:');
  console.log('1. Make sure your FastAPI backend is running');
  console.log('2. Press F5 in VS Code to start debugging');
  console.log('3. Test with: "Hello, how are you?" or upload a file');
  
  console.log('\n🎉 Setup complete! Happy coding!');
  
  rl.close();
}

setup().catch((error) => {
  console.error('❌ Setup failed:', error);
  rl.close();
  process.exit(1);
});