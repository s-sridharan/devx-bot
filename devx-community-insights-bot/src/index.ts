// Entry point using ChatPrompt with custom OpenAIModel
import { config } from 'dotenv';
import restify from 'restify';
import fetch from 'node-fetch';
import { ChatPrompt, IChatModel, Message, ModelMessage } from '@microsoft/teams.ai';
import { McpClientPlugin } from '@microsoft/teams.mcpclient';
import {
  ActivityTypes,
  TurnContext
} from 'botbuilder';


config({ path: '../.env' });


class MinimalOpenAIModel implements IChatModel {
  constructor(private apiKey: string, private model: string) {}

  async send(input: Message): Promise<ModelMessage> {
    const res = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${this.apiKey}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: this.model,
        messages: [{ role: input.role, content: input.content }]
      })
    });

    const json = await res.json();
    return {
      role: 'model',
      content: json.choices[0].message.content
    };
  }
}

const model = new MinimalOpenAIModel(
  process.env.AZURE_OPENAI_KEY ?? '',
  process.env.AZURE_OPENAI_MODEL ?? 'gpt-4o'
);

const prompt = new ChatPrompt<
  Record<string, unknown>,
  [McpClientPlugin]
>(
  {
    instructions: 'You are a helpful assistant analyzing developer insights.',
    model
  },
  [
    new McpClientPlugin({ name: 'mcpClient' })  // ✅ This one is for ChatPrompt
  ]
);


prompt.usePlugin('mcpClient', {
  url: 'https://expert-space-tribble-j6q9vp5p993jwj-3978.app.github.dev/mcp',
  params: {
    headers: {
      'x-functions-key': process.env.AZURE_FUNCTION_KEY ?? ''
    }
  }
});

// import {
//   ConfigurationBotFrameworkAuthentication,
//   ConfigurationServiceClientCredentialFactory
// } from 'botbuilder';

// const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
//   MicrosoftAppId: process.env.BOT_ID,
//   MicrosoftAppPassword: process.env.BOT_PASSWORD,
//   MicrosoftAppType: 'MultiTenant' // or 'SingleTenant' based on your setup
// });

// const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication({}, credentialsFactory);

// const adapter = new CloudAdapter(botFrameworkAuthentication);

const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log('✅ ChatPrompt with custom OpenAI model is running at http://localhost:3978');
});

// Echo POST for Bot Framework Test in Web Chat

server.post('/api/messages', async (req, res) => {
  const context = {
    activity: req.body
  } as TurnContext;

  console.log('Received activity:', context.activity);

  if (context.activity.type === ActivityTypes.Message) {
    const userText = context.activity.text;
    const result = await prompt.send(userText);
    res.send(200, {
      type: 'message',
      text: `[echo] ${result.content}`
    });
  } else {
    res.send(200, {
      type: 'message',
      text: `[system] Received activity of type: ${context.activity.type}`
    });
  }
});

server.post('/chat', async (req, res) => {
  const userInput = req.body.text ?? '';
  const result = await prompt.send(userInput);
  res.send(200, { reply: result.content });
});

// Healthcheck route
server.get('/', (_req, res, next) => {
  res.send(200, '✅ DevX Community Insights Bot is alive.');
  next();
});




