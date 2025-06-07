// src/tools/StackOverflowTool.ts - Standalone MCP tool for Stack Overflow ingestion
/* eslint-disable @typescript-eslint/no-explicit-any */

export interface StackOverflowToolConfig {
  apiKey?: string;
  maxResults?: number;
  defaultTags?: string;
  rateLimitDelay?: number;
}

export interface StackOverflowItem {
  id: string;
  title: string;
  score: number;
  answer_count: number;
  view_count: number;
  tags: string[];
  url: string;
  created_date: string;
  last_activity: string;
  is_answered: boolean;
  excerpt?: string;
  owner?: {
    display_name: string;
    reputation: number;
  };
}

export interface StackOverflowResult {
  tool: 'stackoverflow_ingestion';
  query: string;
  timeframe?: string;
  tags?: string;
  total_results: number;
  quota_remaining: number;
  has_more: boolean;
  execution_time_ms: number;
  api_version: string;
  items: StackOverflowItem[];
  metadata: {
    search_url: string;
    rate_limit_remaining: number;
    ingested_at: string;
    success: boolean;
    error?: string;
  };
}

export class StackOverflowTool {
  private config: StackOverflowToolConfig;
  private baseUrl = 'https://api.stackexchange.com/2.3';
  
  constructor(config: StackOverflowToolConfig = {}) {
    this.config = {
      maxResults: 20,
      defaultTags: 'microsoft-teams,botframework',
      rateLimitDelay: 100,
      ...config
    };
  }

  /**
   * MCP Tool Definition - defines the tool schema for MCP protocol
   */
  static getToolDefinition() {
    return {
      name: 'ingest_stackoverflow_feedback',
      description: 'Ingest developer feedback from Stack Overflow with advanced filtering and metadata extraction',
      inputSchema: {
        type: 'object',
        properties: {
          query: {
            type: 'string',
            description: 'Search query for Stack Overflow (e.g., "teams sdk authentication", "bot framework error")',
            minLength: 3,
            maxLength: 200
          },
          tags: {
            type: 'string',
            description: 'Comma-separated tags to filter results (e.g., "microsoft-teams,botframework,typescript")',
            default: 'microsoft-teams,botframework'
          },
          timeframe: {
            type: 'string',
            enum: ['day', 'week', 'month', 'quarter', 'year', 'all'],
            description: 'Time range for filtering results',
            default: 'month'
          },
          sort: {
            type: 'string',
            enum: ['relevance', 'activity', 'votes', 'creation'],
            description: 'Sort order for results',
            default: 'activity'
          },
          min_score: {
            type: 'number',
            description: 'Minimum score threshold for questions',
            minimum: -10,
            maximum: 1000,
            default: 0
          },
          include_answered: {
            type: 'boolean',
            description: 'Include questions that already have accepted answers',
            default: true
          }
        },
        required: ['query'],
        additionalProperties: false
      }
    };
  }

  /**
   * Execute the Stack Overflow ingestion tool
   */
  async execute(params: {
    query: string;
    tags?: string;
    timeframe?: string;
    sort?: string;
    min_score?: number;
    include_answered?: boolean;
  }): Promise<StackOverflowResult> {
    const startTime = Date.now();
    const ingestedAt = new Date().toISOString();
    
    try {
      console.log(`🔍 StackOverflow Tool: Executing search for "${params.query}"`);
      
      // Validate inputs
      this.validateParams(params);
      
      // Build search URL with all parameters
      const searchUrl = this.buildSearchUrl(params);
      
      // Add rate limiting if configured
      if (this.config.rateLimitDelay && this.config.rateLimitDelay > 0) {
        await this.delay(this.config.rateLimitDelay);
      }
      
      // Execute API call
      const response = await fetch(searchUrl, {
        headers: {
          'User-Agent': 'Community-Insights-Bot/1.0 (Microsoft Teams Platform Analysis)',
          'Accept': 'application/json',
          'Accept-Encoding': 'gzip'
        }
      });

      if (!response.ok) {
        throw new Error(`Stack Overflow API returned ${response.status}: ${response.statusText}`);
      }

      const data = await response.json() as any;
      
      // Check for API errors
      if (data.error_id) {
        throw new Error(`Stack Overflow API Error ${data.error_id}: ${data.error_message}`);
      }

      // Transform raw API data to structured result
      const result = this.transformApiResponse(data, params, searchUrl, startTime, ingestedAt);
      
      console.log(`✅ StackOverflow Tool: Found ${result.total_results} items (${result.execution_time_ms}ms)`);
      
      return result;
      
    } catch (error: any) {
      console.error('❌ StackOverflow Tool Error:', error);
      
      return {
        tool: 'stackoverflow_ingestion',
        query: params.query,
        timeframe: params.timeframe,
        tags: params.tags,
        total_results: 0,
        quota_remaining: 0,
        has_more: false,
        execution_time_ms: Date.now() - startTime,
        api_version: '2.3',
        items: [],
        metadata: {
          search_url: '',
          rate_limit_remaining: 0,
          ingested_at,
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
    
    if (params.min_score !== undefined && (params.min_score < -10 || params.min_score > 1000)) {
      throw new Error('min_score must be between -10 and 1000');
    }
  }

  /**
   * Build the complete search URL with all parameters
   */
  private buildSearchUrl(params: any): string {
    const baseParams = new URLSearchParams({
      order: 'desc',
      sort: params.sort || 'activity',
      intitle: params.query,
      site: 'stackoverflow',
      pagesize: this.config.maxResults?.toString() || '20',
      filter: 'withbody' // Include question body for better analysis
    });

    // Add tags filter
    const tags = params.tags || this.config.defaultTags;
    if (tags) {
      baseParams.set('tagged', tags);
    }

    // Add timeframe filter
    if (params.timeframe && params.timeframe !== 'all') {
      const fromDate = this.calculateFromDate(params.timeframe);
      baseParams.set('fromdate', Math.floor(fromDate.getTime() / 1000).toString());
    }

    // Add score filter
    if (params.min_score !== undefined && params.min_score > 0) {
      baseParams.set('min', params.min_score.toString());
    }

    // Add answered filter
    if (params.include_answered === false) {
      baseParams.set('answers', '0');
    }

    // Add API key if available
    if (this.config.apiKey) {
      baseParams.set('key', this.config.apiKey);
    }

    return `${this.baseUrl}/search?${baseParams.toString()}`;
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
        return now; // Return current date for 'all' or invalid timeframe
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
    ingestedAt: string
  ): StackOverflowResult {
    const items: StackOverflowItem[] = (data.items || []).map((item: any) => ({
      id: `so_${item.question_id}`,
      title: item.title || 'No title',
      score: item.score || 0,
      answer_count: item.answer_count || 0,
      view_count: item.view_count || 0,
      tags: item.tags || [],
      url: item.link || '',
      created_date: item.creation_date ? new Date(item.creation_date * 1000).toISOString() : '',
      last_activity: item.last_activity_date ? new Date(item.last_activity_date * 1000).toISOString() : '',
      is_answered: item.is_answered || false,
      excerpt: item.excerpt || '',
      owner: item.owner ? {
        display_name: item.owner.display_name || 'Anonymous',
        reputation: item.owner.reputation || 0
      } : undefined
    }));

    return {
      tool: 'stackoverflow_ingestion',
      query: params.query,
      timeframe: params.timeframe,
      tags: params.tags,
      total_results: items.length,
      quota_remaining: data.quota_remaining || 0,
      has_more: data.has_more || false,
      execution_time_ms: Date.now() - startTime,
      api_version: '2.3',
      items,
      metadata: {
        search_url: searchUrl,
        rate_limit_remaining: data.quota_remaining || 0,
        ingested_at,
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
  getConfig(): StackOverflowToolConfig {
    return { ...this.config };
  }

  /**
   * Update tool configuration
   */
  updateConfig(newConfig: Partial<StackOverflowToolConfig>): void {
    this.config = { ...this.config, ...newConfig };
  }

  /**
   * Health check - verify API connectivity
   */
  async healthCheck(): Promise<{ healthy: boolean; message: string; response_time_ms?: number }> {
    try {
      const startTime = Date.now();
      const testUrl = `${this.baseUrl}/info?site=stackoverflow`;
      
      const response = await fetch(testUrl, {
        headers: {
          'User-Agent': 'Community-Insights-Bot/1.0 (Health Check)'
        }
      });

      const responseTime = Date.now() - startTime;

      if (response.ok) {
        return {
          healthy: true,
          message: 'Stack Overflow API is accessible',
          response_time_ms: responseTime
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
}

// Export singleton instance with environment configuration
export const stackOverflowTool = new StackOverflowTool({
  apiKey: process.env.STACKOVERFLOW_API_KEY,
  maxResults: parseInt(process.env.STACKOVERFLOW_MAX_RESULTS || '20', 10),
  rateLimitDelay: parseInt(process.env.STACKOVERFLOW_RATE_LIMIT_MS || '100', 10)
});