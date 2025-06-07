// Minimal working solution focused on assignment requirements
import { config } from 'dotenv';
import restify from 'restify';
import { ChatPrompt, IChatModel, Message, ModelMessage } from '@microsoft/teams.ai';
import { McpClientPlugin } from '@microsoft/teams.mcpclient';
import {
  ActivityTypes,
  TurnContext
} from 'botbuilder';

/* eslint-disable @typescript-eslint/no-explicit-any */

config();

config();

console.log('🔍 Azure OpenAI Configuration:');
console.log('Endpoint:', process.env.AZURE_OPENAI_ENDPOINT);
console.log('Model:', process.env.AZURE_OPENAI_MODEL);
console.log('API Version:', process.env.AZURE_OPENAI_API_VERSION);
console.log('API Key:', process.env.AZURE_OPENAI_KEY ? 'SET' : 'MISSING');

class MinimalOpenAIModel implements IChatModel {
  constructor(private apiKey: string, private model: string, private endpoint?: string) {}

  async send(input: Message): Promise<ModelMessage> {
    // Use your actual Azure OpenAI configuration
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

    console.log('🤖 Request body:', JSON.stringify(body, null, 2));

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
    console.log('🤖 Azure OpenAI Response:', JSON.stringify(json, null, 2));

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


// MCP Tools Implementation (Core Assignment Requirement)
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

// Create model and prompt (based on your working code)
const model = new MinimalOpenAIModel(
  process.env.AZURE_OPENAI_KEY || '',  // Use || instead of ??
  process.env.AZURE_OPENAI_MODEL || 'gpt-4o-mini',
  process.env.AZURE_OPENAI_ENDPOINT || ''  // Add default empty string
);

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

RESPONSE FORMAT for Community Analysis:
When analyzing community feedback, ALWAYS structure responses as:

1. **Executive Summary** (2-3 sentences for Ops leadership)
2. **Key Metrics** (quantified insights: # of issues, sentiment trends, affected components)
3. **Priority Pain Points** (top 3-5 issues with category/priority classification)
4. **Actionable Recommendations** (specific next steps for Ops team)
5. **Monitoring Suggestions** (what to track going forward)

TEAMS PLATFORM COMPONENTS EXPERTISE:
- Microsoft Teams SDK (JS/TS, C#, Python)
- Bot Framework v4 components and middleware
- Graph API permissions and scopes
- Teams App manifest and capabilities
- Adaptive Cards schema and templating
- Teams authentication flows (SSO, OAuth)
- App Studio/Developer Portal workflows
- Teams Store submission and compliance

ANALYSIS DEPTH:
- Extract specific error messages and stack traces when available
- Identify patterns across multiple feedback sources
- Correlate issues with recent platform updates or announcements
- Flag breaking changes or deprecated features causing friction
- Assess impact on different developer personas (beginner vs expert)

OUTPUT QUALITY STANDARDS:
- Use precise technical terminology
- Quantify impact where possible (affected developers, frequency)
- Provide timeline estimates for addressing issues
- Reference specific Teams platform documentation when relevant
- Maintain professional tone suitable for Ops team briefings

Remember: Your insights directly inform Microsoft Teams Platform roadmap priorities and developer experience improvements.`,
    model
  },
  [
    new McpClientPlugin({ name: 'mcpClient' })
  ]
);

// const prompt = new ChatPrompt<
//   Record<string, unknown>,
//   [McpClientPlugin]
// >(
//   {
//     instructions: `You are a Community Insights assistant for Microsoft Teams Platform Ops team.

// Your mission: Help Ops team proactively identify and prioritize developer pain points from community forums.

// You have access to MCP tools for:
// 1. Ingesting Stack Overflow feedback 
// 2. Ingesting GitHub issues
// 3. Extracting pain points with AI classification

// When users request community insights:
// 1. Use MCP tools to gather real feedback data
// 2. Extract and classify pain points 
// 3. Provide actionable recommendations for Ops team

// Focus on Teams Platform components: Bot Framework, Teams SDK, Graph API, Adaptive Cards.`,
//     model
//   },
//   [
//     new McpClientPlugin({ name: 'mcpClient' })
//   ]
// );

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

// Enhanced message processing for assignment demo
async function processInsightRequest(userText: string): Promise<string> {
  const text = userText.toLowerCase();
  
  try {
    // Stack Overflow insights
    if (text.includes('stackoverflow') || text.includes('stack overflow')) {
      const queryMatch = userText.match(/for\s+(.+?)(?:\s|$)/i);
      const query = queryMatch ? queryMatch[1] : 'teams platform';
      
      const data = await ingestStackOverflowFeedback(query, 'microsoft-teams,botframework');
      
      if (data.feedback_items.length === 0) {
        return `📊 **Stack Overflow Community Insights for "${query}"**\n\n❌ No recent feedback found. Try broader terms.`;
      }
      
      // Extract pain points from top feedback
      const painPoints = [];
      for (const item of data.feedback_items.slice(0, 3)) {
        const extraction = await extractPainPoints(item.title, 'stackoverflow', model);
        painPoints.push(extraction);
      }
      
      const categoryBreakdown = painPoints.reduce((acc: any, pp) => {
        acc[pp.category] = (acc[pp.category] || 0) + 1;
        return acc;
      }, {});
      
      const priorityBreakdown = painPoints.reduce((acc: any, pp) => {
        acc[pp.priority] = (acc[pp.priority] || 0) + 1;
        return acc;
      }, {});
      
      return `📊 **Stack Overflow Community Insights for "${query}"**

**📈 Summary:**
• Total Discussions: ${data.total_results}
• API Quota Remaining: ${data.quota_remaining}

**🔍 Pain Point Analysis:**
• Categories: ${Object.entries(categoryBreakdown).map(([cat, count]) => `${cat} (${count})`).join(', ')}
• Priority Levels: ${Object.entries(priorityBreakdown).map(([pri, count]) => `${pri} (${count})`).join(', ')}

**⚠️ Key Pain Points for Ops Team:**
${painPoints.map((pp, i) => 
  `${i + 1}. **${pp.primary_pain_point}**
   • Category: ${pp.category} | Priority: ${pp.priority}
   • Sentiment: ${pp.sentiment}
   • Action: ${pp.actionable_insights[0]}
   • Components: ${pp.affected_components.join(', ')}`
).join('\n\n')}

**📋 Top Community Questions:**
${data.feedback_items.slice(0, 3).map((item: any) => 
  `• [${item.title}](${item.url}) (Score: ${item.score}, ${item.answer_count} answers)`
).join('\n')}

**🎯 Ops Recommendations:**
• Focus on ${Object.keys(categoryBreakdown)[0]} issues
• Address ${Object.keys(priorityBreakdown)[0]} priority items first
• Monitor sentiment trends in community feedback`;
    }
    
    // GitHub insights
    else if (text.includes('github')) {
      const queryMatch = userText.match(/for\s+(.+?)(?:\s|$)/i);
      const query = queryMatch ? queryMatch[1] : 'teams platform';
      
      const data = await ingestGitHubFeedback(query, 'microsoft/teams-ai,microsoft/botframework');
      
      if (data.feedback_items.length === 0) {
        return `📊 **GitHub Community Insights for "${query}"**\n\n❌ No recent issues found. Try broader terms.`;
      }
      
      // Extract pain points from issues
      const painPoints = [];
      for (const item of data.feedback_items.slice(0, 3)) {
        const feedbackText = `${item.title} ${item.content.substring(0, 300)}`;
        const extraction = await extractPainPoints(feedbackText, 'github', model);
        painPoints.push(extraction);
      }
      
      const openIssues = data.feedback_items.filter((item: any) => item.state === 'open').length;
      const highPriorityCount = painPoints.filter(pp => pp.priority === 'escalation_risk').length;
      
      return `📊 **GitHub Community Insights for "${query}"**

**📈 Summary:**
• Total Issues Found: ${data.total_results}
• Open Issues: ${openIssues}/${data.feedback_items.length}
• Rate Limit Remaining: ${data.rate_limit_remaining}

**🚨 Critical Insights for Ops Team:**
• High Priority Issues: ${highPriorityCount}
• Escalation Risk Items: ${painPoints.filter(pp => pp.priority === 'escalation_risk').length}

**⚠️ Key Pain Points:**
${painPoints.map((pp, i) => 
  `${i + 1}. **${pp.primary_pain_point}**
   • Repository: ${data.feedback_items[i].repository}
   • Priority: ${pp.priority} | Category: ${pp.category}
   • Action Needed: ${pp.actionable_insights[0]}`
).join('\n\n')}

**📋 Recent Issues:**
${data.feedback_items.slice(0, 3).map((item: any) => 
  `• [${item.title}](${item.url}) - ${item.repository}
    Labels: ${item.labels.join(', ')} | Comments: ${item.comments_count}`
).join('\n')}

**🎯 Ops Action Items:**
• Prioritize ${painPoints.filter(pp => pp.priority === 'escalation_risk').length} escalation-risk issues
• Address ${painPoints.filter(pp => pp.category === 'documentation').length} documentation gaps
• Monitor ${painPoints.filter(pp => pp.sentiment === 'negative').length} negative sentiment items`;
    }
    
    // Direct pain point extraction
    else if (text.includes('extract') || text.includes('classify')) {
      const feedbackMatch = userText.match(/["'](.+?)["']/i) || userText.match(/:\s*(.+)$/i);
      if (feedbackMatch) {
        const feedback = feedbackMatch[1];
        const extraction = await extractPainPoints(feedback, 'user_input', model);
        
        return `🤖 **Pain Point Analysis for Ops Team**

**📝 Feedback:** "${extraction.original_text}"

**🔍 Analysis Results:**
• **Primary Pain Point:** ${extraction.primary_pain_point}
• **Category:** ${extraction.category}
• **Priority Level:** ${extraction.priority}
• **Sentiment:** ${extraction.sentiment}

**🎯 Actionable Insights:**
${extraction.actionable_insights.map((insight: string) => `• ${insight}`).join('\n')}

**🔧 Affected Components:**
${extraction.affected_components.map((comp: string) => `• ${comp}`).join('\n')}

**📊 Ops Recommendation:**
${extraction.priority === 'escalation_risk' ? '🚨 **IMMEDIATE ACTION REQUIRED** - This feedback indicates high escalation risk.' :
  extraction.priority === 'sentiment_spike' ? '📈 **MONITOR CLOSELY** - Sentiment spike detected in community.' :
  '📋 **STANDARD PROCESS** - Add to backlog for regular review.'}`;
      }
    }
    
    // Default AI response
    const result = await prompt.send(userText);
    return result.content || '';
    
  } catch (error: any) {
    console.error('Processing error:', error);
    return `❌ **Error Processing Community Insights Request**

Error: ${error.message}

**Try these MCP-powered commands:**
• "Analyze Stack Overflow feedback for Teams SDK authentication"
• "Find GitHub issues for Bot Framework problems"  
• "Extract pain points from: [paste feedback text here]"
• "Get community insights for Teams Adaptive Cards"`;
  }
}

const port = process.env.PORT || process.env.port || 3978;

server.listen(port, () => {
  console.log(`✅ Community Insights Bot (Assignment Demo) running on http://localhost:${port}`);
  console.log('🔧 MCP Client: Ready for external server connections');
  console.log('📊 Built-in MCP Tools: Stack Overflow ingestion, GitHub ingestion, Pain point extraction');
});

// Main bot endpoint (like your original)
server.post('/api/messages', async (req, res) => {
  console.log('🔍 DEBUG: Received request body:', JSON.stringify(req.body, null, 2));
  
  const context = {
    activity: req.body
  } as TurnContext;

  console.log('📨 Received activity:', context.activity.type, context.activity.text?.substring(0, 50));

  if (context.activity.type === ActivityTypes.Message) {
    try {
      const userText = context.activity.text || '';
      const response = await processInsightRequest(userText);
      
      res.send(200, {
        type: 'message',
        text: response
      });
    } catch (error: any) {
      console.error('❌ Error processing message:', error);
      res.send(200, {
        type: 'message',
        text: '❌ Sorry, I encountered an error processing your community insights request. Please try again.'
      });
    }
  } else {
    res.send(200, {
      type: 'message',
      text: `🤖 Community Insights Bot received activity type: ${context.activity.type}`
    });
  }
});

// MCP Server endpoint (assignment requirement)
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

// Health check
server.get('/', (_req, res, next) => {
  res.send(200, '✅ Community Insights Bot - Assignment Demo Ready!');
  next();
});

server.get('/mcp/status', (_req, res, next) => {
  res.send(200, {
    assignment: 'Community Insights Teams Application',
    mcp_server_status: 'active',
    mcp_client_status: 'ready',
    tools_available: ['ingest_stackoverflow_feedback', 'ingest_github_feedback', 'extract_pain_points'],
    features: ['Feedback Ingestion via MCP', 'AI-Driven Pain Point Extraction', 'Teams Bot Interface']
  });
  next();
});

// // Entry point using ChatPrompt with custom OpenAIModel
// import { config } from 'dotenv';
// import restify from 'restify';
// import fetch from 'node-fetch';
// import { ChatPrompt, IChatModel, Message, ModelMessage } from '@microsoft/teams.ai';
// import { McpClientPlugin } from '@microsoft/teams.mcpclient';
// import {
//   ActivityTypes,
//   TurnContext
// } from 'botbuilder';

// config({ path: '../.env' });

// class MinimalOpenAIModel implements IChatModel {
//   constructor(private apiKey: string, private model: string) {}

//   async send(input: Message): Promise<ModelMessage> {
//     const res = await fetch('https://api.openai.com/v1/chat/completions', {
//       method: 'POST',
//       headers: {
//         'Authorization': `Bearer ${this.apiKey}`,
//         'Content-Type': 'application/json'
//       },
//       body: JSON.stringify({
//         model: this.model,
//         messages: [{ role: input.role, content: input.content }]
//       })
//     });

//     const json = await res.json();
//     return {
//       role: 'model',
//       content: json.choices[0].message.content
//     };
//   }
// }

// const model = new MinimalOpenAIModel(
//   process.env.AZURE_OPENAI_KEY ?? '',
//   process.env.AZURE_OPENAI_MODEL ?? 'gpt-4o-mini'
// );

// const prompt = new ChatPrompt<
//   Record<string, unknown>,
//   [McpClientPlugin]
// >(
//   {
//     instructions: 'You are a helpful assistant analyzing developer insights.',
//     model
//   },
//   [
//     new McpClientPlugin({ name: 'mcpClient' })
//   ]
// )
// .usePlugin('mcpClient', {
//   url: 'https://expert-space-tribble-j6q9vp5p993jwj-9000.app.github.dev/mcp',
//   params: {
//     headers: {
//       'x-functions-key': process.env.AZURE_FUNCTION_KEY ?? ''
//     }
//   }
// });

// const server = restify.createServer();
// server.use(restify.plugins.bodyParser());

// server.listen(process.env.port || process.env.PORT || 3978, () => {
//   console.log('✅ DevX Community Insights Bot is alive on http://localhost:3978');
// });

// server.post('/api/messages', async (req, res) => {
//   const context = {
//     activity: req.body
//   } as TurnContext;

//   console.log('Received activity:', context.activity);

//   if (context.activity.type === ActivityTypes.Message) {
//     const userText = context.activity.text;
//     const result = await prompt.send(userText);
//     res.send(200, {
//       type: 'message',
//       text: `[echo] ${result.content}`
//     });
//   } else {
//     res.send(200, {
//       type: 'message',
//       text: `[system] Received activity of type: ${context.activity.type}`
//     });
//   }
// });

// server.post('/chat', async (req, res) => {
//   const userInput = req.body.text ?? '';
//   const result = await prompt.send(userInput);
//   res.send(200, { reply: result.content });
// });

// server.get('/', (_req, res, next) => {
//   res.send(200, '✅ DevX Community Insights Bot is alive.');
//   next();
// });
