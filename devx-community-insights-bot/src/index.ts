
// // Enhanced Community Insights Bot with full agent chain integration
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */

import { Application, TurnState, TeamsAdapter } from '@microsoft/teams-ai'; // Add this line

import { config } from 'dotenv';
import restify from 'restify';
import { ChatPrompt, IChatModel, Message, ModelMessage } from '@microsoft/teams.ai';
import { McpClientPlugin } from '@microsoft/teams.mcpclient';
import {
  ActivityTypes,
  TurnContext,
  MessageFactory,
  CardFactory,
  ConfigurationServiceClientCredentialFactory, // Add this
  MemoryStorage // Add this
} from 'botbuilder';

// Import all new components
import { ExtractionAgent } from './agents/ExtractionAgent';
import { ResponseAgent } from './agents/ResponseAgent';
import { AgentChain } from './orchestration/AgentChain';
import { ProgressCard } from './cards/ProgressCard';
import { InsightsCard } from './cards/InsightsCard';
import { ActionHandler } from './cards/ActionHandler';
import { GitHubTool } from './tools/GitHubTool';
import { StackOverflowTool } from './tools/StackOverflowTool';
import { 
  RawFeedbackItem, 
  AgentExecutionContext,
  ProgressUpdate,
  AgentChainResult 
} from './types/index';

/* eslint-disable @typescript-eslint/no-explicit-any */

config();

console.log('🔍 Azure OpenAI Configuration:');
console.log('Endpoint:', process.env.AZURE_OPENAI_ENDPOINT);
console.log('Model:', process.env.AZURE_OPENAI_MODEL);
console.log('API Version:', process.env.AZURE_OPENAI_API_VERSION);
console.log('API Key:', process.env.AZURE_OPENAI_KEY ? 'SET' : 'MISSING');

// ✅ ADD THIS DEBUG CODE temporarily
console.log('🔐 DEBUG Bot Credentials:');
console.log('BOT_ID from env:', process.env.BOT_ID);
console.log('BOT_PASSWORD from env:', process.env.BOT_PASSWORD ? 'SET (length: ' + process.env.BOT_PASSWORD.length + ')' : 'MISSING');
console.log('BOT_TENANT_ID from env:', process.env.BOT_TENANT_ID);
console.log('Current working directory:', process.cwd());

// Keep your working MinimalOpenAIModel implementation
class MinimalOpenAIModel implements IChatModel {
  constructor(private apiKey: string, private model: string, private endpoint?: string) {}

  async send(input: Message): Promise<ModelMessage> {
    const url = `${this.endpoint}/openai/deployments/${this.model}/chat/completions?api-version=${process.env.AZURE_OPENAI_API_VERSION || '2024-12-01-preview'}`;
    
    console.log('🤖 Calling Azure OpenAI URL:', url);
    
    const headers = {
      'api-key': this.apiKey,
      'Content-Type': 'application/json'
    };

    const body = {
      messages: [
        {
          role: "system",
          content: "You are a helpful assistant that analyzes developer feedback and extracts pain points."
        },
        {
          role: input.role,
          content: input.content
        }
      ],
      max_tokens: 1000,
      temperature: 0.3,
      top_p: 1.0
    };

    const res = await fetch(url, {
      method: 'POST',
      headers,
      body: JSON.stringify(body)
    });

    if (!res.ok) {
      const errorText = await res.text();
      console.error('❌ Azure OpenAI Error:', res.status, errorText);
      throw new Error(`Azure OpenAI API error: ${res.status} - ${errorText}`);
    }

    const json = await res.json() as any;

    if (!json.choices || !json.choices[0] || !json.choices[0].message) {
      console.error('❌ Invalid Azure OpenAI response:', json);
      throw new Error('Invalid response from Azure OpenAI');
    }

    return {
      role: 'model',
      content: json.choices[0].message.content || 'No response generated'
    };
  }
}

// Initialize model and components
const model = new MinimalOpenAIModel(
  process.env.AZURE_OPENAI_KEY || '',
  process.env.AZURE_OPENAI_MODEL || 'gpt-4o-mini',
  process.env.AZURE_OPENAI_ENDPOINT || ''
);

// Initialize all new components
const githubTool = new GitHubTool({
  token: process.env.GITHUB_TOKEN,
  maxResults: 20,
  defaultRepositories: 'microsoft/teams-ai,microsoft/botframework'
});

const stackOverflowTool = new StackOverflowTool({
  apiKey: process.env.STACKOVERFLOW_API_KEY,
  maxResults: 20,
  defaultTags: 'microsoft-teams,botframework'
});

const agentChain = new AgentChain(model);
const progressCard = new ProgressCard({ show_technical_details: true });
const insightsCard = new InsightsCard({ show_technical_metrics: true });
const actionHandler = new ActionHandler(model, githubTool, stackOverflowTool);

// ✅ ADD THIS: Teams AI adapter and application setup
const adapter = new TeamsAdapter(
    {},
    new ConfigurationServiceClientCredentialFactory({
        MicrosoftAppId: process.env.BOT_ID,
        MicrosoftAppPassword: process.env.BOT_PASSWORD,
        MicrosoftAppTenantId: process.env.BOT_TENANT_ID,
        MicrosoftAppType: 'SingleTenant'
    })
);

const onTurnErrorHandler = async (context: TurnContext, error: any) => {
    console.error(`\n [onTurnError] unhandled error: ${error}`);
    await context.sendActivity('The bot encountered an error or bug.');
};

adapter.onTurnError = onTurnErrorHandler;

interface ConversationState {
    count: number;
}
type ApplicationTurnState = TurnState<ConversationState>;

const storage = new MemoryStorage();
const app = new Application<ApplicationTurnState>({
    storage
});

// ✅ ADD THIS: Message handler using your existing logic
app.activity(ActivityTypes.Message, async (context: TurnContext, state: ApplicationTurnState) => {
    const userText = context.activity.text?.trim();
    
    if (userText?.toLowerCase().includes('analyze')) {
        try {
            console.log('🎯 Processing analysis request:', userText);
            const resultCard = await processInsightRequestForWebChat(userText);
            
            await context.sendActivity({
                attachments: [{
                    contentType: 'application/vnd.microsoft.card.adaptive',
                    content: resultCard
                }]
            });
            return;
        } catch (error: any) {
            console.error('❌ Analysis error:', error);
            await context.sendActivity(`❌ Error: ${error.message}`);
            return;
        }
    }

    // Default echo for other messages
    let count = state.conversation.count ?? 0;
    state.conversation.count = ++count;
    await context.sendActivity(`[${count}] you said: ${userText}`);
});

// Keep your MCP prompt setup
const prompt = new ChatPrompt<
  Record<string, unknown>,
  [McpClientPlugin]
>(
  {
    instructions: `You are the Community Insights AI Engine for Microsoft Teams Platform Operations team.

MISSION: Help Ops team proactively identify and prioritize developer pain points from community feedback to improve Teams Platform developer experience.

CONTEXT: You analyze feedback from Stack Overflow and GitHub to surface actionable insights for:
- Bot Framework SDK issues
- Teams JavaScript/TypeScript SDK problems  
- Graph API integration challenges
- Adaptive Cards implementation issues
- Authentication and SSO difficulties
- Teams App deployment and distribution problems

PAIN POINT CLASSIFICATION FRAMEWORK:
Category Classification:
- "documentation": Missing/unclear docs, insufficient examples, outdated guides
- "engineering": Bugs, performance issues, SDK limitations, API reliability
- "product": Feature gaps, UX friction, platform limitations, developer tool needs

Priority Classification:
- "escalation_risk": Multiple negative mentions, blocking issues, security concerns
- "sentiment_spike": Sudden increase in complaints, viral negative feedback
- "normal": Standard feedback, feature requests, general questions

Remember: Your insights directly inform Microsoft Teams Platform roadmap priorities and developer experience improvements.`,
    model
  },
  [
    new McpClientPlugin({ name: 'mcpClient' })
  ]
);

// Add your external MCP server if available
if (process.env.EXTERNAL_MCP_URL) {
  prompt.usePlugin('mcpClient', {
    url: process.env.EXTERNAL_MCP_URL,
    params: {
      headers: {
        'x-functions-key': process.env.AZURE_FUNCTION_KEY ?? ''
      }
    }
  });
}

const server = restify.createServer();
server.use(restify.plugins.bodyParser());

// Legacy MCP tools (keep for compatibility)
async function ingestStackOverflowFeedback(query: string, tags?: string) {
  console.log(`🔍 MCP Tool: Ingesting Stack Overflow feedback for "${query}"`);
  
  let url = `https://api.stackexchange.com/2.3/search?order=desc&sort=activity&intitle=${encodeURIComponent(query)}&site=stackoverflow&pagesize=10`;
  
  if (tags) {
    url += `&tagged=${encodeURIComponent(tags)}`;
  }
  
  if (process.env.STACKOVERFLOW_API_KEY) {
    url += `&key=${process.env.STACKOVERFLOW_API_KEY}`;
  }

  const response = await fetch(url);
  const data = await response.json() as any;
  
  return {
    tool: 'ingest_stackoverflow_feedback',
    source: 'stackoverflow',
    query,
    total_results: data.items?.length || 0,
    quota_remaining: data.quota_remaining,
    feedback_items: data.items?.slice(0, 5).map((item: any) => ({
      id: `so_${item.question_id}`,
      title: item.title,
      score: item.score,
      answer_count: item.answer_count,
      view_count: item.view_count,
      tags: item.tags || [],
      url: item.link,
      created_date: new Date(item.creation_date * 1000).toISOString(),
      is_answered: item.is_answered
    })) || []
  };
}

async function ingestGitHubFeedback(query: string, repositories?: string) {
  console.log(`🔍 MCP Tool: Ingesting GitHub feedback for "${query}"`);
  
  let searchQuery = `${query} is:issue`;
  if (repositories) {
    const repos = repositories.split(',');
    searchQuery += ` repo:${repos.join(' repo:')}`;
  }

  const url = `https://api.github.com/search/issues?q=${encodeURIComponent(searchQuery)}&sort=updated&per_page=10`;
  
  const headers: any = {
    'Accept': 'application/vnd.github.v3+json',
    'User-Agent': 'Community-Insights-Bot'
  };
  
  if (process.env.GITHUB_TOKEN) {
    headers['Authorization'] = `token ${process.env.GITHUB_TOKEN}`;
  }

  const response = await fetch(url, { headers });
  const data = await response.json() as any;
  
  return {
    tool: 'ingest_github_feedback',
    source: 'github',
    query,
    total_results: data.total_count || 0,
    rate_limit_remaining: response.headers.get('x-ratelimit-remaining'),
    feedback_items: data.items?.slice(0, 5).map((item: any) => ({
      id: `gh_${item.id}`,
      title: item.title,
      content: item.body || '',
      state: item.state,
      repository: item.repository_url.split('/').slice(-2).join('/'),
      url: item.html_url,
      labels: item.labels?.map((l: any) => l.name) || [],
      author: item.user?.login || 'Anonymous',
      created_date: item.created_at,
      comments_count: item.comments
    })) || []
  };
}

async function postGitHubComment(repository: string, issueNumber: number, commentBody: string) {
  console.log(`🔧 MCP Tool: Posting comment to ${repository}#${issueNumber}`);
  
  const [owner, repo] = repository.split('/');
  const url = `https://api.github.com/repos/${owner}/${repo}/issues/${issueNumber}/comments`;
  
  // ✅ FIXED: Current GitHub API format (2024)
  const headers = {
    'Accept': 'application/vnd.github+json',              // ← FIXED
    'Authorization': `Bearer ${process.env.GITHUB_TOKEN}`, // ← FIXED: Bearer instead of token
    'X-GitHub-Api-Version': '2022-11-28',                // ← ADDED: Required
    'User-Agent': 'Community-Insights-Bot',
    'Content-Type': 'application/json'
  };
  
  const response = await fetch(url, {
    method: 'POST',
    headers,
    body: JSON.stringify({ body: commentBody })
  });
  
  if (!response.ok) {
    const error = await response.text();
    console.error('❌ GitHub API error:', response.status, error);
    return {
      tool: 'post_github_comment',
      success: false,
      error: `GitHub API error: ${response.status} - ${error}`
    };
  }
  
  const comment = await response.json() as any;
  
  console.log('✅ GitHub comment posted:', comment.id);
  
  return {
    tool: 'post_github_comment',
    success: true,
    comment_id: comment.id,
    comment_url: comment.html_url,
    repository: repository,
    issue_number: issueNumber,
    created_at: comment.created_at
  };
}

async function extractPainPoints(feedbackText: string, source: string, model: MinimalOpenAIModel) {
  console.log(`🤖 MCP Tool: Extracting pain points from ${source} feedback`);
  
  const prompt = `Analyze this developer feedback and extract pain points for Teams Platform Ops team:

FEEDBACK: "${feedbackText}"
SOURCE: ${source}

Return JSON with:
{
  "primary_pain_point": "main issue",
  "category": "documentation|engineering|product", 
  "priority": "escalation_risk|sentiment_spike|normal",
  "sentiment": "positive|negative|neutral",
  "actionable_insights": ["insight1", "insight2"],
  "affected_components": ["component1", "component2"]
}`;

  const result = await model.send({ role: 'user', content: prompt });
  
  try {
    const extraction = result.content ? JSON.parse(result.content) : 
      {
        primary_pain_point: 'No content returned',
        category: 'engineering',
        priority: 'normal',
        sentiment: 'neutral',
        actionable_insights: ['Manual review needed'],
        affected_components: ['unknown']
      };
    return {
      tool: 'extract_pain_points',
      ...extraction,
      source,
      original_text: feedbackText.substring(0, 200),
      extracted_at: new Date().toISOString()
    };
  } catch {
    return {
      tool: 'extract_pain_points',
      primary_pain_point: 'Failed to parse',
      category: 'engineering',
      priority: 'normal',
      sentiment: 'neutral',
      actionable_insights: ['Manual review needed'],
      affected_components: ['unknown'],
      source,
      original_text: feedbackText.substring(0, 200),
      extracted_at: new Date().toISOString()
    };
  }
}

const port = process.env.PORT || process.env.port || 3978;

server.listen(port, () => {
  console.log(`✅ Enhanced Community Insights Bot running on http://localhost:${port}`);
  console.log('🔧 MCP Client: Ready for external server connections');
  console.log('🤖 Agent Chain: ExtractionAgent → ResponseAgent ready');
  console.log('🎯 Teams Cards: Progress, Insights, and Action cards enabled');
  console.log('📊 Tools: Enhanced GitHub + Stack Overflow ingestion');
});

// Full processInsightRequestForWebChat function with debugging
async function processInsightRequestForWebChat(userText: string): Promise<any> {
  const text = userText.toLowerCase();
  const startTime = Date.now();
  const requestId = `req_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  
  let platform: 'stackoverflow' | 'github' = 'stackoverflow';
  let query = 'teams platform';
  
  // Determine platform and query
  if (text.includes('github')) {
    platform = 'github';
    const queryMatch = userText.match(/for\s+(.+?)(?:\s|$)/i);
    query = queryMatch ? queryMatch[1] : 'teams platform';
  } else if (text.includes('stackoverflow') || text.includes('stack overflow')) {
    platform = 'stackoverflow';
    const queryMatch = userText.match(/for\s+(.+?)(?:\s|$)/i);
    query = queryMatch ? queryMatch[1] : 'teams sdk authentication';
  }

  console.log(`🎯 Processing: "${query}" on ${platform}`);

  // Create execution context
  const executionContext: AgentExecutionContext = {
    request_id: requestId,
    user_query: userText,
    timestamp: new Date().toISOString(),
    debug_mode: process.env.NODE_ENV !== 'production'
  };

  // Progress callback (just logs for Web Chat)
  const progressCallback = async (update: ProgressUpdate) => {
    console.log(`📊 ${update.stage}: ${update.progress_percentage}% - ${update.message}`);
  };

  // Gather feedback data using your existing tools
  let rawFeedback: RawFeedbackItem[] = [];
  
  if (platform === 'stackoverflow') {
    const soResult = await stackOverflowTool.execute({
      query: query,
      tags: 'microsoft-teams,botframework,typescript',
      timeframe: 'month',
      sort: 'activity'
    });
    
    rawFeedback = soResult.items.map((item: any) => ({
      id: item.id,
      title: item.title,
      content: item.excerpt || '',
      url: item.url,
      source: 'stackoverflow' as const,
      metadata: {
        score: item.score,
        tags: item.tags,
        author: item.owner?.display_name,
        created_date: item.created_date
      }
    }));
  } else {
    const ghResult = await githubTool.execute({
      query: query,
      repositories: 'microsoft/teams-ai,microsoft/botframework',
      timeframe: 'month',
      state: 'open'
    });
    
    rawFeedback = ghResult.items.map((item: any) => ({
      id: item.id,
      title: item.title,
      content: item.body,
      url: item.url,
      source: 'github' as const,
      metadata: {
        score: item.reactions.total_count,
        tags: item.labels,
        author: item.author,
        created_date: item.created_date,
        repository: item.repository
      }
    }));
  }

  if (rawFeedback.length === 0) {
    console.log('❌ No feedback found, returning error card');
    // Return error card using your existing component
    const errorCard = insightsCard.generateErrorCard({
      stage: 'ingestion',
      message: `No recent feedback found for "${query}" on ${platform}`,
      technical_details: 'Try broader search terms or different time range'
    }, {
      user_query: query,
      platform_analyzed: platform
    });
    
    console.log('📤 Error card generated:', !!errorCard.card);
    return errorCard.card;
  }

  // Execute agent chain using your existing orchestrator
  const chainResult: AgentChainResult = await agentChain.execute({
    raw_feedback: rawFeedback,
    target_context: {
      platform: platform === 'github' ? 'github' : 'general',
      audience: 'developer_community',
      urgency_level: 'normal'
    },
    response_requirements: {
      include_technical_details: true,
      include_workarounds: true,
      include_timelines: true,
      max_length: 2000
    },
    execution_context: executionContext,
    progress_callback: progressCallback
  });

  console.log('🎯 AgentChain result:', {
    success: chainResult.overall_success,
    painPoints: chainResult.extraction_result?.data?.pain_points?.length,
    responseGenerated: !!chainResult.response_result?.data?.community_response,
    executionTime: chainResult.total_execution_time_ms
  });

  try {
    // Generate insights card using your existing component
    const cardResult = insightsCard.generateCard(chainResult, {
      user_query: query,
      platform_analyzed: platform,
      execution_time_ms: chainResult.total_execution_time_ms
    });

    console.log('✅ Card generated successfully:', {
      cardType: cardResult.card?.type,
      bodyItems: cardResult.card?.body?.length,
      actions: cardResult.card?.actions?.length,
      hasCard: !!cardResult.card
    });

    // Return the card for Web Chat
    return cardResult.card;

  } catch (cardError: any) {
    console.error('❌ Card generation failed:', cardError);
    
    // Return simple fallback card
    const fallbackCard = {
      type: 'AdaptiveCard',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: '🎯 Analysis Complete!',
          weight: 'bolder',
          size: 'large'
        },
        {
          type: 'TextBlock',
          text: `Found ${chainResult.extraction_result?.data?.pain_points?.length || 0} pain points from your query: "${userText}"`,
          wrap: true
        },
        {
          type: 'TextBlock',
          text: chainResult.response_result?.data?.community_response?.response_text || 'Response generated successfully!',
          wrap: true,
          size: 'small'
        },
        {
          type: 'FactSet',
          facts: [
            { title: 'Platform', value: platform },
            { title: 'Query', value: query },
            { title: 'Execution Time', value: `${chainResult.total_execution_time_ms}ms` },
            { title: 'Pain Points', value: (chainResult.extraction_result?.data?.pain_points?.length || 0).toString() }
          ]
        }
      ]
    };

    console.log('📤 Fallback card generated');
    return fallbackCard;
  }
}

server.post('/api/messages', async (req, res) => {
    console.log('📨 API Messages: Using Teams AI pattern');
    await adapter.process(req, res as any, async (context) => {
        await analysisApp.run(context);
    });
});


const analysisStorage = new MemoryStorage();
const analysisApp = new Application<ApplicationTurnState>({
    storage: analysisStorage
});

// ✅ ANALYSIS MESSAGE HANDLER (replaces your current logic):
analysisApp.activity(ActivityTypes.Message, async (context: TurnContext, state: ApplicationTurnState) => {
    const userText = context.activity.text?.trim();

        // ✅ HANDLE CARD ACTIONS (when text is undefined)
    if (!userText && context.activity.value) {
        console.log('🎯 Card action detected:', context.activity.value);
        
        const actionData = context.activity.value;
        if (actionData.action === 'post_community_response') {
            await context.sendActivity('🔄 Posting to GitHub...');
            
            try {
                const result = await postToGitHubViaMCP(
                    actionData.repository,
                    actionData.issue_number,
                    actionData.response_text
                );
                
                if (result.success) {
                    await context.sendActivity(`✅ Posted! Comment ID: ${result.comment_id}`);
                } else {
                    await context.sendActivity(`❌ Failed: ${result.error}`);
                }
            } catch (error: any) {
                await context.sendActivity(`❌ Error: ${error.message}`);
            }
        }
        return;
    }
    
    if (userText?.toLowerCase().includes('analyze')) {
        // ✅ IMMEDIATE PROGRESS CARD
        const progressCard = {
            type: 'AdaptiveCard',
            version: '1.2',
            body: [
                {
                    type: 'Container',
                    style: 'emphasis',
                    items: [
                        {
                            type: 'TextBlock',
                            text: '🔄 Analysis Starting',
                            weight: 'Bolder',
                            size: 'Large'
                        }
                    ]
                },
                {
                    type: 'TextBlock',
                    text: `Processing: "${userText}"`,
                    wrap: true
                },
                {
                    type: 'TextBlock',
                    text: '📊 Live updates coming...',
                    isSubtle: true
                }
            ]
        };
        
        await context.sendActivity({
            attachments: [{
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: progressCard
            }]
        });
        
        // ✅ START ANALYSIS WITH LIVE UPDATES
        await processInsightRequest(userText, context);
        
    } else {
        // Default echo
        await context.sendActivity(`Echo: ${userText}`);
    }
});

// ✅ SIMPLE: Handle card button clicks
analysisApp.activity('invoke', async (context: TurnContext, state: ApplicationTurnState) => {
    const invokeValue = context.activity.value;
    console.log('🎯 Card action:', invokeValue?.action);
    
    if (invokeValue?.action === 'post_community_response') {
        await context.sendActivity('🔄 Posting to GitHub...');
        
        try {
            const result = await postToGitHubViaMCP(
                invokeValue.repository,
                invokeValue.issue_number, 
                invokeValue.response_text
            );
            
            if (result.success) {
                await context.sendActivity(`✅ Posted! Comment ID: ${result.comment_id}`);
            } else {
                await context.sendActivity(`❌ Failed: ${result.error}`);
            }
        } catch (error: any) {
            await context.sendActivity(`❌ Error: ${error.message}`);
        }
    }
});


// ✅ SIMPLE: Call your MCP server  
async function postToGitHubViaMCP(repository: string, issueNumber: number, responseText: string) {
    const response = await fetch('http://localhost:3978/mcp', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            method: 'tools/call',
            params: {
                name: 'post_github_comment',
                arguments: { repository, issue_number: issueNumber, comment_body: responseText }
            }
        })
    });
    
    const result = await response.json();
    return JSON.parse(result.content[0].text);
}

// ✅ UPDATE your existing processInsightRequest function to send live updates:
async function processInsightRequest(userText: string, context: TurnContext): Promise<void> {
  const text = userText.toLowerCase();
  const startTime = Date.now();
  const requestId = `req_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    
  // Determine platform and query
  let platform: 'stackoverflow' | 'github' = 'stackoverflow';
  let query = 'teams platform';
  
  try {
    if (text.includes('github')) {
      platform = 'github';
      const queryMatch = userText.match(/for\s+(.+?)(?:\s|$)/i);
      query = queryMatch ? queryMatch[1] : 'teams platform';
    } else if (text.includes('stackoverflow') || text.includes('stack overflow')) {
      platform = 'stackoverflow';
      const queryMatch = userText.match(/for\s+(.+?)(?:\s|$)/i);
      query = queryMatch ? queryMatch[1] : 'teams sdk authentication';
    }

    // Create execution context
    const executionContext = {
      request_id: requestId,
      user_query: userText,
      timestamp: new Date().toISOString(),
      debug_mode: process.env.NODE_ENV !== 'production'
    };

    // ✅ LIVE PROGRESS UPDATES - send messages to user
    const progressCallback = async (update: any) => {
      await context.sendActivity(`📊 ${update.stage}: ${update.progress_percentage}% - ${update.message}`);
    };

    // Start with progress update
    await progressCallback({
      stage: 'ingestion',
      message: `Analyzing ${platform} feedback for: "${query}"`,
      progress_percentage: 10,
      current_operation: 'Gathering community feedback data'
    });

    // Gather feedback data
    let rawFeedback: any[] = [];
    
    if (platform === 'stackoverflow') {
      const soResult = await stackOverflowTool.execute({
        query: query,
        tags: 'microsoft-teams,botframework,typescript',
        timeframe: 'month',
        sort: 'activity'
      });
      
      rawFeedback = soResult.items.map((item: any) => ({
        id: item.id,
        title: item.title,
        content: item.excerpt || '',
        url: item.url,
        source: 'stackoverflow' as const,
        metadata: {
          score: item.score,
          tags: item.tags,
          author: item.owner?.display_name,
          created_date: item.created_date
        }
      }));
    } else {
      const ghResult = await githubTool.execute({
        query: query,
        repositories: 'microsoft/teams-ai,microsoft/botframework',
        timeframe: 'month',
        state: 'open'
      });
      
      rawFeedback = ghResult.items.map((item: any) => ({
        id: item.id,
        title: item.title,
        content: item.body,
        url: item.url,
        source: 'github' as const,
        metadata: {
          score: item.reactions.total_count,
          tags: item.labels,
          author: item.author,
          created_date: item.created_date,
          repository: item.repository
        }
      }));
    }

    if (rawFeedback.length === 0) {
      await context.sendActivity(`❌ No recent feedback found for "${query}" on ${platform}`);
      return;
    }

    // Execute agent chain with progress tracking
    const chainResult = await agentChain.execute({
      raw_feedback: rawFeedback,
      target_context: {
        platform: platform === 'github' ? 'github' : 'general',
        audience: 'developer_community',
        urgency_level: 'normal'
      },
      response_requirements: {
        include_technical_details: true,
        include_workarounds: true,
        include_timelines: true,
        max_length: 2000
      },
      execution_context: executionContext,
      progress_callback: progressCallback
    });

    // Generate insights card with results
    const cardResult = insightsCard.generateCard(chainResult, {
      user_query: query,
      platform_analyzed: platform,
      execution_time_ms: chainResult.total_execution_time_ms
    });

    await context.sendActivity({
        attachments: [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: cardResult.card
        }]
    });

  } catch (error: any) {
    console.error('❌ Error processing insight request:', error);
    await context.sendActivity(`❌ Error: ${error.message}`);
  }
}


// NEW: Action handler endpoint for Teams card actions
server.post('/api/actions', async (req, res) => {
  console.log('🎯 Received card action:', req.body.action);
  
  try {
    const actionData = req.body;
    const actionContext = {
      user_id: actionData.from?.id || 'unknown',
      conversation_id: actionData.conversation?.id || 'unknown',
      original_query: actionData.original_query || 'unknown',
      action_timestamp: new Date().toISOString()
    };

    // Create mock context for action handler
    const context = {
      activity: req.body,
      sendActivity: async (activity: any) => {
        console.log('📤 Action response:', activity.attachments?.[0]?.contentType || activity.text?.substring(0, 50));
        return { id: `action_response_${Date.now()}` };
      }
    } as TurnContext;

    const result = await actionHandler.handleAction(context, actionData, actionContext);
    
    if (result.success) {
      res.send(200, {
        type: 'message',
        text: result.message,
        attachments: result.card ? [CardFactory.adaptiveCard(result.card)] : undefined
      });
    } else {
      res.send(200, {
        type: 'message',
        text: `❌ Action failed: ${result.message}`,
        attachments: result.card ? [CardFactory.adaptiveCard(result.card)] : undefined
      });
    }
  } catch (error: any) {
    console.error('❌ Action handler error:', error);
    res.send(200, {
      type: 'message',
      text: '❌ Sorry, the action failed. Please try again.'
    });
  }
});

// Keep your MCP server endpoint
server.post('/mcp', async (req, res) => {
  try {
    const { method, params } = req.body;
    
    if (method === 'tools/list') {
      res.send(200, {
        tools: [
          {
            name: 'ingest_stackoverflow_feedback',
            description: 'Ingest developer feedback from Stack Overflow for Teams Platform analysis',
            inputSchema: {
              type: 'object',
              properties: {
                query: { type: 'string', description: 'Search query for Stack Overflow' },
                tags: { type: 'string', description: 'Optional tags filter' }
              },
              required: ['query']
            }
          },
          {
            name: 'ingest_github_feedback',
            description: 'Ingest developer feedback from GitHub issues for Teams Platform analysis',
            inputSchema: {
              type: 'object',
              properties: {
                query: { type: 'string', description: 'Search query for GitHub issues' },
                repositories: { type: 'string', description: 'Optional repository filter' }
              },
              required: ['query']
            }
          },
          {
            name: 'extract_pain_points',
            description: 'Extract and classify pain points from developer feedback using AI',
            inputSchema: {
              type: 'object',
              properties: {
                feedback_text: { type: 'string', description: 'Feedback text to analyze' },
                source: { type: 'string', description: 'Source platform' }
              },
              required: ['feedback_text', 'source']
            }
          },
          {
          name: 'post_github_comment',
          description: 'Post a comment to a GitHub issue',
          inputSchema: {
            type: 'object',
            properties: {
              repository: { type: 'string', description: 'Repository in format owner/repo' },
              issue_number: { type: 'number', description: 'Issue number to comment on' },
              comment_body: { type: 'string', description: 'Comment text to post' }
            },
            required: ['repository', 'issue_number', 'comment_body']
          }
        }
        ]
      });
    } else if (method === 'tools/call') {
      const { name, arguments: args } = params;
      let result;
      
      switch (name) {
        case 'ingest_stackoverflow_feedback':
          result = await ingestStackOverflowFeedback(args.query, args.tags);
          break;
        case 'ingest_github_feedback':
          result = await ingestGitHubFeedback(args.query, args.repositories);
          break;
        case 'extract_pain_points':
          result = await extractPainPoints(args.feedback_text, args.source, model);
          break;
        case 'post_github_comment':
          result = await postGitHubComment(args.repository, args.issue_number, args.comment_body);
          break;
        default:
          throw new Error(`Unknown MCP tool: ${name}`);
      }
      
      res.send(200, { 
        content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] 
      });
    } else {
      res.send(400, { error: 'Unknown MCP method' });
    }
  } catch (error: any) {
    console.error('MCP endpoint error:', error);
    res.send(500, { error: error.message });
  }
});

// Keep your health check endpoints
server.get('/', (_req, res, next) => {
  res.send(200, '✅ Enhanced Community Insights Bot - Production Ready!');
  next();
});

server.get('/mcp/status', (_req, res, next) => {
  res.send(200, {
    assignment: 'Community Insights Teams Application',
    mcp_server_status: 'active',
    mcp_client_status: 'ready',
    agent_chain_status: 'ready',
    tools_available: ['Enhanced GitHub Ingestion', 'Enhanced Stack Overflow Ingestion', 'AI Pain Point Extraction'],
    features: [
      'Agent Chain Processing (ExtractionAgent → ResponseAgent)',
      'Real-time Progress Cards',
      'Interactive Insights Cards',
      'GitHub Action Integration',
      'Professional Community Response Generation',
      'MCP Protocol Support'
    ],
    endpoints: {
      messages: '/api/messages',
      actions: '/api/actions',
      mcp: '/mcp',
      status: '/mcp/status'
    }
  });
  next();
});

// NEW: Health check for all components
server.get('/health', async (_req, res) => {
  try {
    const githubHealth = await githubTool.healthCheck();
    const stackOverflowHealth = await stackOverflowTool.healthCheck();
    const chainHealth = await agentChain.healthCheck();

    res.send(200, {
      status: 'healthy',
      timestamp: new Date().toISOString(),
      components: {
        github_tool: githubHealth,
        stackoverflow_tool: stackOverflowHealth,
        agent_chain: chainHealth,
        azure_openai: {
          configured: !!process.env.AZURE_OPENAI_KEY,
          endpoint: !!process.env.AZURE_OPENAI_ENDPOINT
        }
      }
    });
  } catch (error: any) {
    res.send(500, {
      status: 'unhealthy',
      error: error.message,
      timestamp: new Date().toISOString()
    });
  }
});
