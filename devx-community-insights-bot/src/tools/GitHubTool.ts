// src/tools/GitHubTool.ts - Standalone MCP tool for GitHub Issues ingestion
/* eslint-disable @typescript-eslint/no-explicit-any */

export interface GitHubToolConfig {
  token?: string;
  maxResults?: number;
  defaultRepositories?: string;
  rateLimitDelay?: number;
  userAgent?: string;
}

export interface GitHubIssueItem {
  id: string;
  title: string;
  body: string;
  number: number;
  state: 'open' | 'closed';
  repository: string;
  url: string;
  created_date: string;
  updated_date: string;
  labels: string[];
  author: string;
  assignees: string[];
  comments_count: number;
  reactions: {
    total_count: number;
    thumbs_up: number;
    thumbs_down: number;
    laugh: number;
    hooray: number;
    confused: number;
    heart: number;
    rocket: number;
    eyes: number;
  };
  milestone?: {
    title: string;
    state: string;
    due_on?: string;
  };
}

export interface GitHubResult {
  tool: 'github_ingestion';
  query: string;
  timeframe?: string;
  repositories?: string;
  total_results: number;
  incomplete_results: boolean;
  rate_limit_remaining: number;
  execution_time_ms: number;
  api_version: string;
  items: GitHubIssueItem[];
  metadata: {
    search_url: string;
    rate_limit_reset: string;
    ingested_at: string;
    success: boolean;
    error?: string;
  };
}

export class GitHubTool {
  private config: GitHubToolConfig;
  private baseUrl = 'https://api.github.com';
  
  constructor(config: GitHubToolConfig = {}) {
    this.config = {
      maxResults: 20,
      defaultRepositories: 'microsoft/teams-ai,microsoft/botframework',
      rateLimitDelay: 100,
      userAgent: 'Community-Insights-Bot/1.0',
      ...config
    };
  }

  /**
   * MCP Tool Definition - defines the tool schema for MCP protocol
   */
  static getToolDefinition() {
    return {
      name: 'ingest_github_feedback',
      description: 'Ingest developer feedback from GitHub Issues with advanced filtering and metadata extraction',
      inputSchema: {
        type: 'object',
        properties: {
          query: {
            type: 'string',
            description: 'Search query for GitHub issues (e.g., "authentication bug", "bot framework error", "teams sdk documentation")',
            minLength: 3,
            maxLength: 200
          },
          repositories: {
            type: 'string',
            description: 'Comma-separated repository names to search (e.g., "microsoft/teams-ai,microsoft/botframework")',
            default: 'microsoft/teams-ai,microsoft/botframework'
          },
          timeframe: {
            type: 'string',
            enum: ['day', 'week', 'month', 'quarter', 'year', 'all'],
            description: 'Time range for filtering results',
            default: 'month'
          },
          state: {
            type: 'string',
            enum: ['open', 'closed', 'all'],
            description: 'Issue state filter',
            default: 'open'
          },
          labels: {
            type: 'string',
            description: 'Comma-separated labels to filter (e.g., "bug,enhancement,documentation")'
          },
          sort: {
            type: 'string',
            enum: ['created', 'updated', 'comments', 'reactions'],
            description: 'Sort order for results',
            default: 'updated'
          },
          include_pull_requests: {
            type: 'boolean',
            description: 'Include pull requests in results',
            default: false
          }
        },
        required: ['query'],
        additionalProperties: false
      }
    };
  }

  /**
   * Execute the GitHub ingestion tool
   */
  async execute(params: {
    query: string;
    repositories?: string;
    timeframe?: string;
    state?: string;
    labels?: string;
    sort?: string;
    include_pull_requests?: boolean;
  }): Promise<GitHubResult> {
    const startTime = Date.now();
    const ingestedAt = new Date().toISOString();
    
    try {
      console.log(`🔍 GitHub Tool: Executing search for "${params.query}"`);
      
      // Validate inputs
      this.validateParams(params);
      
      // Build search URL with all parameters
      const searchUrl = this.buildSearchUrl(params);
      
      // Add rate limiting if configured
      if (this.config.rateLimitDelay && this.config.rateLimitDelay > 0) {
        await this.delay(this.config.rateLimitDelay);
      }
      
      // Execute API call
      const headers: Record<string, string> = {
        'Accept': 'application/vnd.github.v3+json',
        'User-Agent': this.config.userAgent || 'Community-Insights-Bot/1.0'
      };
      
      if (this.config.token) {
        headers['Authorization'] = `token ${this.config.token}`;
      }

      const response = await fetch(searchUrl, { headers });

      if (!response.ok) {
        throw new Error(`GitHub API returned ${response.status}: ${response.statusText}`);
      }

      const data = await response.json() as any;
      
      // Check for API errors
      if (data.message) {
        throw new Error(`GitHub API Error: ${data.message}`);
      }

      // Transform raw API data to structured result
      const result = this.transformApiResponse(data, params, searchUrl, startTime, ingestedAt, response);
      
      console.log(`✅ GitHub Tool: Found ${result.total_results} items (${result.execution_time_ms}ms)`);
      
      return result;
      
    } catch (error: any) {
      console.error('❌ GitHub Tool Error:', error);
      
      return {
        tool: 'github_ingestion',
        query: params.query,
        timeframe: params.timeframe,
        repositories: params.repositories,
        total_results: 0,
        incomplete_results: true,
        rate_limit_remaining: 0,
        execution_time_ms: Date.now() - startTime,
        api_version: 'v3',
        items: [],
        metadata: {
          search_url: '',
          rate_limit_reset: '',
          ingested_at: ingestedAt,
          success: false,
          error: error.message
        }
      };
    }
  }

  /**
   * Validate input parameters
   */
  private validateParams(params: any): void {
    if (!params.query || params.query.trim().length < 3) {
      throw new Error('Query must be at least 3 characters long');
    }
    
    if (params.query.length > 200) {
      throw new Error('Query must be 200 characters or less');
    }
    
    // Validate repository format if provided
    if (params.repositories) {
      const repos = params.repositories.split(',');
      for (const repo of repos) {
        if (!repo.trim().includes('/')) {
          throw new Error(`Invalid repository format: "${repo}". Use format "owner/repo"`);
        }
      }
    }
  }

  /**
   * Build the complete search URL with all parameters
   */
  private buildSearchUrl(params: any): string {
    // Start with base query
    let searchQuery = params.query;
    
    // Add issue type filter (exclude PRs by default)
    if (!params.include_pull_requests) {
      searchQuery += ' is:issue';
    }
    
    // Add repository filter
    const repositories = params.repositories || this.config.defaultRepositories;
    if (repositories) {
    const repos = repositories.split(',').map((r: string) => r.trim());  // ✅ Fixed: added (r: string)
    searchQuery += ` repo:${repos.join(' repo:')}`;
    }
    
    // Add state filter
    if (params.state && params.state !== 'all') {
      searchQuery += ` state:${params.state}`;
    }
    
    // Add timeframe filter
    if (params.timeframe && params.timeframe !== 'all') {
      const fromDate = this.calculateFromDate(params.timeframe);
      searchQuery += ` created:>=${fromDate.toISOString().split('T')[0]}`;
    }
    
    // Add labels filter
    if (params.labels) {
    const labels = params.labels.split(',').map((l: string) => l.trim());  // ✅ Fixed: added (l: string)
    for (const label of labels) {
        searchQuery += ` label:"${label}"`;
    }
    }

    // Build URL parameters
    const urlParams = new URLSearchParams({
      q: searchQuery,
      sort: params.sort || 'updated',
      order: 'desc',
      per_page: this.config.maxResults?.toString() || '20'
    });

    return `${this.baseUrl}/search/issues?${urlParams.toString()}`;
  }

  /**
   * Calculate from date based on timeframe
   */
  private calculateFromDate(timeframe: string): Date {
    const now = new Date();
    const fromDate = new Date();
    
    switch (timeframe) {
      case 'day':
        fromDate.setDate(now.getDate() - 1);
        break;
      case 'week':
        fromDate.setDate(now.getDate() - 7);
        break;
      case 'month':
        fromDate.setMonth(now.getMonth() - 1);
        break;
      case 'quarter':
        fromDate.setMonth(now.getMonth() - 3);
        break;
      case 'year':
        fromDate.setFullYear(now.getFullYear() - 1);
        break;
      default:
        return now;
    }
    
    return fromDate;
  }

  /**
   * Transform raw API response to structured result
   */
  private transformApiResponse(
    data: any, 
    params: any, 
    searchUrl: string, 
    startTime: number, 
    ingestedAt: string,
    response: Response
  ): GitHubResult {
    const items: GitHubIssueItem[] = (data.items || []).map((item: any) => ({
      id: `gh_${item.id}`,
      title: item.title || 'No title',
      body: item.body || '',
      number: item.number || 0,
      state: item.state || 'open',
      repository: item.repository_url ? item.repository_url.split('/').slice(-2).join('/') : 'unknown',
      url: item.html_url || '',
      created_date: item.created_at || '',
      updated_date: item.updated_at || '',
      labels: (item.labels || []).map((label: any) => label.name),
      author: item.user?.login || 'Anonymous',
      assignees: (item.assignees || []).map((assignee: any) => assignee.login),
      comments_count: item.comments || 0,
      reactions: {
        total_count: item.reactions?.total_count || 0,
        thumbs_up: item.reactions?.['+1'] || 0,
        thumbs_down: item.reactions?.['-1'] || 0,
        laugh: item.reactions?.laugh || 0,
        hooray: item.reactions?.hooray || 0,
        confused: item.reactions?.confused || 0,
        heart: item.reactions?.heart || 0,
        rocket: item.reactions?.rocket || 0,
        eyes: item.reactions?.eyes || 0
      },
      milestone: item.milestone ? {
        title: item.milestone.title,
        state: item.milestone.state,
        due_on: item.milestone.due_on
      } : undefined
    }));

    // Extract rate limit info from headers
    const rateLimitRemaining = parseInt(response.headers.get('x-ratelimit-remaining') || '0', 10);
    const rateLimitReset = response.headers.get('x-ratelimit-reset') || '';

    return {
      tool: 'github_ingestion',
      query: params.query,
      timeframe: params.timeframe,
      repositories: params.repositories,
      total_results: data.total_count || items.length,
      incomplete_results: data.incomplete_results || false,
      rate_limit_remaining: rateLimitRemaining,
      execution_time_ms: Date.now() - startTime,
      api_version: 'v3',
      items,
      metadata: {
        search_url: searchUrl,
        rate_limit_reset: rateLimitReset,
        ingested_at: ingestedAt,
        success: true
      }
    };
  }

  /**
   * Simple delay utility for rate limiting
   */
  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Get tool configuration for debugging/monitoring
   */
  getConfig(): GitHubToolConfig {
    return { ...this.config };
  }

  /**
   * Update tool configuration
   */
  updateConfig(newConfig: Partial<GitHubToolConfig>): void {
    this.config = { ...this.config, ...newConfig };
  }

  /**
   * Health check - verify API connectivity and rate limits
   */
  async healthCheck(): Promise<{ healthy: boolean; message: string; response_time_ms?: number; rate_limit_remaining?: number }> {
    try {
      const startTime = Date.now();
      const testUrl = `${this.baseUrl}/rate_limit`;
      
      const headers: Record<string, string> = {
        'User-Agent': this.config.userAgent || 'Community-Insights-Bot/1.0'
      };
      
      if (this.config.token) {
        headers['Authorization'] = `token ${this.config.token}`;
      }

      const response = await fetch(testUrl, { headers });
      const responseTime = Date.now() - startTime;

      if (response.ok) {
        const data = await response.json() as any;
        const remaining = data.rate?.remaining || 0;
        
        return {
          healthy: true,
          message: `GitHub API is accessible. Rate limit: ${remaining} requests remaining`,
          response_time_ms: responseTime,
          rate_limit_remaining: remaining
        };
      } else {
        return {
          healthy: false,
          message: `API returned ${response.status}: ${response.statusText}`
        };
      }
    } catch (error: any) {
      return {
        healthy: false,
        message: `Health check failed: ${error.message}`
      };
    }
  }

  /**
   * Get available repositories for search suggestions
   */
  getDefaultRepositories(): string[] {
    return (this.config.defaultRepositories || '').split(',').map(repo => repo.trim()).filter(Boolean);
  }

  /**
   * Estimate API quota usage for a search
   */
  estimateQuotaUsage(params: any): { estimated_calls: number; quota_type: 'search' | 'core'; notes: string } {
    // GitHub Search API has different rate limits than Core API
    const isComplexSearch = params.repositories && params.repositories.split(',').length > 3;
    const estimated_calls = isComplexSearch ? 2 : 1;
    
    return {
      estimated_calls,
      quota_type: 'search',
      notes: this.config.token 
        ? 'Authenticated requests: 30 searches per minute' 
        : 'Unauthenticated requests: 10 searches per minute'
    };
  }
}

// Export singleton instance with environment configuration
export const gitHubTool = new GitHubTool({
  token: process.env.GITHUB_TOKEN,
  maxResults: parseInt(process.env.GITHUB_MAX_RESULTS || '20', 10),
  rateLimitDelay: parseInt(process.env.GITHUB_RATE_LIMIT_MS || '100', 10),
  defaultRepositories: process.env.GITHUB_DEFAULT_REPOS || 'microsoft/teams-ai,microsoft/botframework'
});