// src/cards/InsightsCard.ts - Generate Teams Adaptive Cards for agent chain results and actions
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */

import { 
  AgentChainResult, 
  InsightsCardData, 
  CardAction,
  ExtractedPainPoint,
  CommunityResponse 
} from '../types/index.js';

export interface InsightsCardOptions {
  show_technical_metrics?: boolean;
  compact_mode?: boolean;
  theme?: 'default' | 'dark' | 'high-contrast';
  max_pain_points_displayed?: number;
  include_confidence_indicators?: boolean;
  enable_quick_actions?: boolean;
}

export interface InsightsCardResult {
  card: any; // Adaptive Card JSON
  actions_available: string[];
  accessibility_text: string;
  estimated_user_engagement: number;
}

export class InsightsCard {
  private options: InsightsCardOptions;

  constructor(options: InsightsCardOptions = {}) {
    this.options = {
      show_technical_metrics: true,
      compact_mode: false,
      theme: 'default',
      max_pain_points_displayed: 5,
      include_confidence_indicators: true,
      enable_quick_actions: true,
      ...options
    };
  }

  /**
   * Generate insights card from agent chain results
   */
  generateCard(
    chainResult: AgentChainResult,
    context: {
      user_query: string;
      platform_analyzed: string;
      execution_time_ms: number;
    }
  ): InsightsCardResult {

    const cardData = this.transformChainResultToCardData(chainResult, context);
    const visualStyling = this.determineVisualStyling(chainResult);

    // Build the main card structure
    const card = {
      type: 'AdaptiveCard',
      version: '1.4',
      schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      body: [
        // Header with summary
        this.buildHeaderSection(cardData, context),
        
        // Analysis overview  
        this.buildAnalysisSection(cardData),
        
        // Pain points section
        this.buildPainPointsSection(cardData.priority_insights),
        
        // Generated response section
        ...(cardData.generated_response ? [this.buildResponseSection(cardData.generated_response)] : []),
        
        // Quality metrics (if enabled)
        ...(this.options.show_technical_metrics ? [this.buildMetricsSection(chainResult)] : []),
        
        // Quick insights callout
        this.buildQuickInsightsSection(cardData)
      ],
      actions: this.options.enable_quick_actions ? this.buildCardActions(cardData.actions) : [],
      ...((visualStyling.color_theme === 'attention' || visualStyling.color_theme === 'warning') ? { 
        style: 'attention' 
      } : {})
    };

    return {
      card,
      actions_available: cardData.actions.map(a => this.mapActionType(a.type)),
      accessibility_text: this.buildAccessibilityText(cardData, context),
      estimated_user_engagement: this.estimateUserEngagement(cardData, chainResult)
    };
  }

  /**
   * Generate error card for failed analysis
   */
  generateErrorCard(
    error: {
      stage: string;
      message: string;
      technical_details?: string;
    },
    context: {
      user_query: string;
      platform_analyzed: string;
    }
  ): InsightsCardResult {

    const card = {
      type: 'AdaptiveCard',
      version: '1.4',
      schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      body: [
        // Error header
        {
          type: 'ColumnSet',
          columns: [
            {
              type: 'Column',
              width: 'auto',
              items: [{
                type: 'Image',
                url: 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iNDAiIGhlaWdodD0iNDAiIHZpZXdCb3g9IjAgMCA0MCA0MCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGNpcmNsZSBjeD0iMjAiIGN5PSIyMCIgcj0iMTgiIGZpbGw9IiNGRjQ0NDIiLz4KPHBhdGggZD0iTTIwIDEwVjI0IiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjMiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIvPgo8Y2lyY2xlIGN4PSIyMCIgY3k9IjMwIiByPSIyIiBmaWxsPSJ3aGl0ZSIvPgo8L3N2Zz4K',
                width: '40px',
                height: '40px'
              }]
            },
            {
              type: 'Column',
              width: 'stretch',
              items: [
                {
                  type: 'TextBlock',
                  text: '⚠️ Analysis Failed',
                  weight: 'bolder',
                  size: 'large',
                  color: 'attention'
                },
                {
                  type: 'TextBlock',
                  text: `Unable to analyze feedback for: "${context.user_query}"`,
                  wrap: true,
                  size: 'medium'
                },
                {
                  type: 'TextBlock',
                  text: error.message,
                  wrap: true,
                  size: 'small',
                  color: 'default'
                }
              ]
            }
          ]
        },

        // Error details
        ...(error.technical_details ? [{
          type: 'Container',
          style: 'emphasis',
          items: [
            {
              type: 'TextBlock',
              text: '**Technical Details:**',
              weight: 'bolder',
              size: 'small'
            },
            {
              type: 'TextBlock',
              text: error.technical_details,
              wrap: true,
              size: 'small',
              fontType: 'monospace'
            }
          ]
        }] : []),

        // Alternative actions
        {
          type: 'Container',
          style: 'good',
          items: [{
            type: 'TextBlock',
            text: '💡 **Try these alternatives:**\n• Refine your search query\n• Check Stack Overflow directly\n• Contact our support team',
            wrap: true,
            size: 'small'
          }]
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: '🔄 Try Different Query',
          data: { action: 'retry_with_new_query' },
          style: 'positive'
        },
        {
          type: 'Action.OpenUrl',
          title: '📖 View Documentation',
          url: 'https://docs.microsoft.com/en-us/microsoftteams/platform/'
        },
        {
          type: 'Action.Submit',
          title: '📞 Contact Support',
          data: { action: 'contact_support', context: error.stage }
        }
      ]
    };

    return {
      card,
      actions_available: ['retry_with_new_query', 'contact_support'],
      accessibility_text: `Analysis failed during ${error.stage}. ${error.message}`,
      estimated_user_engagement: 0.3
    };
  }

  /**
   * Transform chain result to card data structure
   */
  private transformChainResultToCardData(
    chainResult: AgentChainResult, 
    context: any
  ): InsightsCardData {
    
    const painPoints = chainResult.extraction_result?.data?.pain_points || [];
    const response = chainResult.response_result?.data?.community_response;
    const suggestedActions = chainResult.response_result?.data?.suggested_actions || [];

    // Transform pain points to priority insights
    const priorityInsights = painPoints
      .slice(0, this.options.max_pain_points_displayed)
      .map(pp => ({
        title: pp.primary_pain_point,
        category: pp.category,
        priority: pp.priority,
        sentiment: pp.sentiment,
        description: this.buildPainPointDescription(pp)
      }));

    // Transform suggested actions to card actions
    const cardActions: CardAction[] = suggestedActions.map(action => ({
      type: this.mapActionType(action.action_type),
      label: action.action_label,
      enabled: true,
      action_data: action.action_data,
      style: action.priority === 'high' ? 'positive' : 'default'
    }));

    return {
      analysis_summary: {
        platform: context.platform_analyzed,
        query: context.user_query,
        total_items: chainResult.extraction_result?.data?.summary_statistics?.total_items_analyzed || 0,
        pain_points_count: painPoints.length
      },
      priority_insights: priorityInsights,
      generated_response: response ? {
        text: response.response_text,
        confidence: response.confidence_level,
        includes_workaround: response.includes_workaround
      } : undefined,
      actions: cardActions,
      visual_styling: this.determineVisualStyling(chainResult)
    };
  }

  /**
   * Build header section with analysis summary
   */
  private buildHeaderSection(cardData: InsightsCardData, context: any) {
    const statusIcon = this.getAnalysisStatusIcon(cardData);
    const summaryText = this.buildSummaryText(cardData);

    return {
      type: 'ColumnSet',
      columns: [
        {
          type: 'Column',
          width: 'auto',
          items: [{
            type: 'Image',
            url: statusIcon.data_url,
            width: '48px',
            height: '48px'
          }]
        },
        {
          type: 'Column',
          width: 'stretch',
          items: [
            {
              type: 'TextBlock',
              text: '🎯 **Community Insights Analysis**',
              weight: 'bolder',
              size: 'large'
            },
            {
              type: 'TextBlock',
              text: summaryText,
              wrap: true,
              size: 'medium',
              color: 'default'
            },
            {
              type: 'TextBlock',
              text: `📊 Query: "${cardData.analysis_summary.query}" • Platform: ${cardData.analysis_summary.platform} • Completed in ${Math.round(context.execution_time_ms / 1000)}s`,
              wrap: true,
              size: 'small',
              color: 'accent'
            }
          ]
        }
      ]
    };
  }

  /**
   * Build analysis overview section
   */
  private buildAnalysisSection(cardData: InsightsCardData) {
    const facts = [
      {
        title: 'Items Analyzed',
        value: cardData.analysis_summary.total_items.toString()
      },
      {
        title: 'Pain Points Found',
        value: cardData.analysis_summary.pain_points_count.toString()
      },
      {
        title: 'Response Generated',
        value: cardData.generated_response ? 'Yes' : 'No'
      },
      {
        title: 'Actions Available',
        value: cardData.actions.length.toString()
      }
    ];

    return {
      type: 'FactSet',
      facts: facts
    };
  }

  /**
   * Build pain points section
   */
  private buildPainPointsSection(insights: any[]) {
    if (insights.length === 0) {
      return {
        type: 'Container',
        style: 'emphasis',
        items: [{
          type: 'TextBlock',
          text: '✅ No critical pain points detected in the analyzed feedback.',
          wrap: true,
          horizontalAlignment: 'center'
        }]
      };
    }

    return {
      type: 'Container',
      items: [
        {
          type: 'TextBlock',
          text: '🔍 **Key Pain Points Identified:**',
          weight: 'bolder',
          size: 'medium'
        },
        ...insights.map((insight, index) => ({
          type: 'Container',
          style: this.getPainPointStyle(insight.priority),
          items: [
            {
              type: 'ColumnSet',
              columns: [
                {
                  type: 'Column',
                  width: 'auto',
                  items: [{
                    type: 'TextBlock',
                    text: this.getPriorityEmoji(insight.priority),
                    size: 'medium'
                  }]
                },
                {
                  type: 'Column',
                  width: 'stretch',
                  items: [
                    {
                      type: 'TextBlock',
                      text: `**${insight.title}**`,
                      weight: 'bolder',
                      size: 'small',
                      wrap: true
                    },
                    {
                      type: 'TextBlock',
                      text: insight.description,
                      wrap: true,
                      size: 'small'
                    },
                    {
                      type: 'TextBlock',
                      text: `Category: ${insight.category} • Priority: ${insight.priority} • Sentiment: ${insight.sentiment}`,
                      size: 'small',
                      color: 'accent',
                      fontType: 'monospace'
                    }
                  ]
                }
              ]
            }
          ]
        }))
      ]
    };
  }

  /**
   * Build response section with generated community response
   */
  private buildResponseSection(response: any) {
    const confidenceColor = response.confidence > 0.8 ? 'good' : 
                           response.confidence > 0.6 ? 'warning' : 'attention';

    return {
      type: 'Container',
      style: 'emphasis',
      items: [
        {
          type: 'TextBlock',
          text: '✍️ **Generated Community Response:**',
          weight: 'bolder',
          size: 'medium'
        },
        {
          type: 'Container',
          style: 'good',
          items: [{
            type: 'TextBlock',
            text: this.truncateText(response.text, 500),
            wrap: true,
            size: 'small'
          }]
        },
        {
          type: 'ColumnSet',
          columns: [
            {
              type: 'Column',
              width: 'stretch',
              items: [{
                type: 'TextBlock',
                text: `Confidence: ${Math.round(response.confidence * 100)}% ${response.includes_workaround ? '• Includes workaround' : ''}`,
                size: 'small',
                color: 'accent'
              }]
            },
            {
              type: 'Column',
              width: 'auto',
              items: [{
                type: 'TextBlock',
                text: this.getConfidenceEmoji(response.confidence),
                size: 'medium'
              }]
            }
          ]
        }
      ]
    };
  }

  /**
   * Build technical metrics section
   */
  private buildMetricsSection(chainResult: AgentChainResult) {
    const extractionMeta = chainResult.extraction_result?.metadata;
    const responseMeta = chainResult.response_result?.metadata;

    const metrics = [
      {
        title: 'Total Execution Time',
        value: `${chainResult.total_execution_time_ms}ms`
      },
      {
        title: 'AI Model Calls',
        value: ((extractionMeta?.model_calls || 0) + (responseMeta?.model_calls || 0)).toString()
      },
      {
        title: 'Quality Score',
        value: `${Math.round((chainResult.chain_metadata?.quality_score || 0) * 100)}%`
      },
      {
        title: 'Steps Completed',
        value: `${chainResult.chain_metadata?.steps_completed || 0}/2`
      }
    ];

    return {
      type: 'Container',
      style: 'emphasis',
      items: [
        {
          type: 'TextBlock',
          text: '📊 **Technical Metrics:**',
          weight: 'bolder',
          size: 'small'
        },
        {
          type: 'FactSet',
          facts: metrics
        }
      ]
    };
  }

  /**
   * Build quick insights callout
   */
  private buildQuickInsightsSection(cardData: InsightsCardData) {
    const insights = this.generateQuickInsights(cardData);

    return {
      type: 'Container',
      style: 'good',
      items: [
        {
          type: 'TextBlock',
          text: '💡 **Quick Insights:**',
          weight: 'bolder',
          size: 'small'
        },
        {
          type: 'TextBlock',
          text: insights.join('\n'),
          wrap: true,
          size: 'small'
        }
      ]
    };
  }

  /**
   * Build card-level actions
   */
  private buildCardActions(actions: CardAction[]) {
    return actions.slice(0, 4).map(action => ({
      type: 'Action.Submit',
      title: action.label,
      data: {
        action: action.type,
        ...action.action_data
      },
      style: action.style || 'default'
    }));
  }

  // Helper methods
  private mapActionType(actionType: string): 'post_response' | 'escalate' | 'track_issue' | 'update_docs' | 'retry_analysis' {
    const actionMap: Record<string, 'post_response' | 'escalate' | 'track_issue' | 'update_docs' | 'retry_analysis'> = {
      'post_response': 'post_response',
      'escalate_internal': 'escalate',
      'update_documentation': 'update_docs',
      'create_tracking_issue': 'track_issue',
      'retry_analysis': 'retry_analysis'
    };
    
    return actionMap[actionType] || 'retry_analysis';
  }

  private buildPainPointDescription(pp: ExtractedPainPoint): string {
    const components = pp.affected_components.length > 0 
      ? ` Affects: ${pp.affected_components.slice(0, 3).join(', ')}`
      : '';
    
    const severity = pp.impact_assessment.blocking_severity !== 'medium' 
      ? ` Severity: ${pp.impact_assessment.blocking_severity}`
      : '';

    return `${pp.primary_pain_point.substring(0, 100)}${components}${severity}`;
  }

  private buildSummaryText(cardData: InsightsCardData): string {
    const { total_items, pain_points_count } = cardData.analysis_summary;
    
    if (pain_points_count === 0) {
      return `Analyzed ${total_items} feedback items - no critical issues found! 🎉`;
    }
    
    return `Found ${pain_points_count} pain points from ${total_items} feedback items. ${cardData.generated_response ? 'Community response ready!' : 'Manual review needed.'}`;
  }

  private getAnalysisStatusIcon(cardData: InsightsCardData) {
    const hasResponse = !!cardData.generated_response;
    const hasCritical = cardData.priority_insights.some(p => p.priority === 'escalation_risk');
    
    if (hasCritical) {
      return {
        data_url: 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iNDgiIGhlaWdodD0iNDgiIHZpZXdCb3g9IjAgMCA0OCA0OCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGNpcmNsZSBjeD0iMjQiIGN5PSIyNCIgcj0iMjIiIGZpbGw9IiNGRjhDMDAiLz4KPHBhdGggZD0iTTI0IDEyVjI4IiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjMiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIvPgo8Y2lyY2xlIGN4PSIyNCIgY3k9IjM0IiByPSIyIiBmaWxsPSJ3aGl0ZSIvPgo8L3N2Zz4K'
      };
    }
    
    if (hasResponse) {
      return {
        data_url: 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iNDgiIGhlaWdodD0iNDgiIHZpZXdCb3g9IjAgMCA0OCA0OCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGNpcmNsZSBjeD0iMjQiIGN5PSIyNCIgcj0iMjIiIGZpbGw9IiMxNkExNzUiLz4KPHBhdGggZD0iTTE2IDI0TDIyIDMwTDMyIDE4IiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjMiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCIvPgo8L3N2Zz4K'
      };
    }
    
    return {
      data_url: 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iNDgiIGhlaWdodD0iNDgiIHZpZXdCb3g9IjAgMCA0OCA0OCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGNpcmNsZSBjeD0iMjQiIGN5PSIyNCIgcj0iMjIiIGZpbGw9IiM0Qzg0RkYiLz4KPHBhdGggZD0iTTE2IDE2SDMyVjMySDI0VjI0SDE2VjE2WiIgZmlsbD0id2hpdGUiLz4KPC9zdmc+Cg=='
    };
  }

  private getPainPointStyle(priority: string): string {
    return priority === 'escalation_risk' ? 'attention' : 
           priority === 'sentiment_spike' ? 'warning' : 'default';
  }

  private getPriorityEmoji(priority: string): string {
    const emojiMap: Record<string, string> = {
      'escalation_risk': '🚨',
      'sentiment_spike': '⚠️',
      'normal': '📝'
    };
    return emojiMap[priority] || '📝';
  }

  private getConfidenceEmoji(confidence: number): string {
    if (confidence >= 0.9) return '🎯';
    if (confidence >= 0.7) return '✅';
    if (confidence >= 0.5) return '⚠️';
    return '❓';
  }

  private generateQuickInsights(cardData: InsightsCardData): string[] {
    const insights = [];
    
    if (cardData.priority_insights.length > 3) {
      insights.push(`• High volume of feedback (${cardData.priority_insights.length} pain points)`);
    }
    
    const criticalCount = cardData.priority_insights.filter(p => p.priority === 'escalation_risk').length;
    if (criticalCount > 0) {
      insights.push(`• ${criticalCount} critical issues requiring immediate attention`);
    }
    
    if (cardData.generated_response && cardData.generated_response.includes_workaround) {
      insights.push('• Response includes practical workarounds for developers');
    }
    
    const categories = [...new Set(cardData.priority_insights.map(p => p.category))];
    if (categories.length === 1) {
      insights.push(`• All issues are ${categories[0]}-related - consider focused initiative`);
    }
    
    if (insights.length === 0) {
      insights.push('• Analysis completed successfully with actionable results');
    }
    
    return insights;
  }

  private determineVisualStyling(chainResult: AgentChainResult) {
    const hasErrors = !chainResult.overall_success;
    const hasCritical = chainResult.extraction_result?.data?.pain_points?.some(
      pp => pp.priority === 'escalation_risk'
    );
    
    return {
      color_theme: hasErrors ? 'warning' as const : 
                   hasCritical ? 'attention' as const : 'success' as const,
      priority_indicators: true,
      progress_bars: false
    };
  }

  private truncateText(text: string, maxLength: number): string {
    return text.length > maxLength ? `${text.substring(0, maxLength)}...` : text;
  }

  private buildAccessibilityText(cardData: InsightsCardData, context: any): string {
    const { total_items, pain_points_count } = cardData.analysis_summary;
    return `Community insights analysis complete. Analyzed ${total_items} feedback items for query "${context.user_query}". Found ${pain_points_count} pain points. ${cardData.generated_response ? 'Community response generated.' : 'Manual review needed.'} ${cardData.actions.length} actions available.`;
  }

  private estimateUserEngagement(cardData: InsightsCardData, chainResult: AgentChainResult): number {
    let engagement = 0.5; // Base engagement
    
    if (cardData.generated_response) engagement += 0.2;
    if (cardData.priority_insights.length > 0) engagement += 0.1;
    if (cardData.actions.length > 0) engagement += 0.1;
    if (chainResult.chain_metadata?.quality_score && chainResult.chain_metadata.quality_score > 0.8) engagement += 0.1;
    
    return Math.min(1.0, engagement);
  }

  /**
   * Update options
   */
  updateOptions(newOptions: Partial<InsightsCardOptions>): void {
    this.options = { ...this.options, ...newOptions };
  }

  /**
   * Get current options  
   */
  getOptions(): InsightsCardOptions {
    return { ...this.options };
  }
}