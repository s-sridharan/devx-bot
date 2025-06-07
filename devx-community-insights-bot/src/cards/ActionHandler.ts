// src/cards/ActionHandler.ts - Handle Teams Adaptive Card action submissions
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */

import { TurnContext, MessageFactory } from 'botbuilder';
import { IChatModel } from '@microsoft/teams.ai';
import { GitHubTool } from '../tools/GitHubTool.js';
import { StackOverflowTool } from '../tools/StackOverflowTool.js';
import { AgentChain } from '../orchestration/AgentChain.js';
import { ProgressCard } from './ProgressCard.js';
import { InsightsCard } from './InsightsCard.js';
import { 
  CommunityResponse,
  AgentExecutionContext,
  RawFeedbackItem
} from '../types/index.js';

export interface ActionHandlerConfig {
  github_demo_issue_url?: string;
  enable_real_posting?: boolean;
  enable_internal_actions?: boolean;
  max_retry_attempts?: number;
}

export interface ActionContext {
  user_id: string;
  conversation_id: string;
  original_query: string;
  action_timestamp: string;
}

export interface ActionResult {
  success: boolean;
  message: string;
  card?: any; // Follow-up Adaptive Card
  next_actions?: string[];
  execution_time_ms: number;
  metadata: {
    action_type: string;
    user_impact: 'high' | 'medium' | 'low';
    requires_follow_up: boolean;
    debug_info?: any;
  };
}

export class ActionHandler {
  private config: ActionHandlerConfig;
  private model: IChatModel;
  private githubTool: GitHubTool;
  private stackOverflowTool: StackOverflowTool;
  private agentChain: AgentChain;
  private progressCard: ProgressCard;
  private insightsCard: InsightsCard;

  constructor(
    model: IChatModel,
    githubTool: GitHubTool,
    stackOverflowTool: StackOverflowTool,
    config: ActionHandlerConfig = {}
  ) {
    this.model = model;
    this.githubTool = githubTool;
    this.stackOverflowTool = stackOverflowTool;
    this.agentChain = new AgentChain(model);
    this.progressCard = new ProgressCard();
    this.insightsCard = new InsightsCard();
    
    this.config = {
      github_demo_issue_url: process.env.DEMO_GITHUB_ISSUE_URL || 'https://api.github.com/repos/s-sridharan/devx-bot/issues/4/comments',
      enable_real_posting: process.env.NODE_ENV === 'production',
      enable_internal_actions: true,
      max_retry_attempts: 2,
      ...config
    };
  }

  /**
   * Handle action submission from Teams Adaptive Card
   */
  async handleAction(
    context: TurnContext,
    actionData: any,
    actionContext: ActionContext
  ): Promise<ActionResult> {
    const startTime = Date.now();
    const actionId = `action_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    
    try {
      console.log(`🎯 ActionHandler: Processing ${actionData.action} for user ${actionContext.user_id}`);
      
      // Validate action data
      this.validateActionData(actionData);
      
      // Route to specific action handler
      const result = await this.routeAction(context, actionData, actionContext);
      
      const executionTime = Date.now() - startTime;
      console.log(`✅ ActionHandler: Completed ${actionData.action} in ${executionTime}ms`);
      
      return {
        ...result,
        execution_time_ms: executionTime,
        metadata: {
          ...result.metadata,
          debug_info: {
            ...result.metadata.debug_info,
            action_id: actionId,
            processing_time: executionTime
          }
        }
      };
      
    } catch (error: any) {
      const executionTime = Date.now() - startTime;
      console.error('❌ ActionHandler Error:', error);
      
      return {
        success: false,
        message: `Action failed: ${error.message}`,
        card: this.createErrorCard(actionData.action, error.message),
        execution_time_ms: executionTime,
        metadata: {
          action_type: actionData.action || 'unknown',
          user_impact: 'high',
          requires_follow_up: true,
          debug_info: {
            action_id: actionId,
            error_type: error.name || 'UnknownError',
            error_location: 'ActionHandler.handleAction'
          }
        }
      };
    }
  }

  /**
   * Route action to specific handler
   */
  private async routeAction(
    context: TurnContext,
    actionData: any,
    actionContext: ActionContext
  ): Promise<Omit<ActionResult, 'execution_time_ms'>> {
    
    switch (actionData.action) {
      case 'post_response':
        return await this.handlePostResponse(context, actionData, actionContext);
        
      case 'escalate':
        return await this.handleEscalateInternal(context, actionData, actionContext);
        
      case 'track_issue':
        return await this.handleCreateTrackingIssue(context, actionData, actionContext);
        
      case 'update_docs':
        return await this.handleUpdateDocumentation(context, actionData, actionContext);
        
      case 'retry_analysis':
        return await this.handleRetryAnalysis(context, actionData, actionContext);
        
      case 'contact_support':
        return await this.handleContactSupport(context, actionData, actionContext);
        
      default:
        throw new Error(`Unknown action type: ${actionData.action}`);
    }
  }

  /**
   * Handle posting community response to GitHub
   */
  private async handlePostResponse(
    context: TurnContext,
    actionData: any,
    actionContext: ActionContext
  ): Promise<Omit<ActionResult, 'execution_time_ms'>> {
    
    try {
      // Extract response text from action data
      const responseText = actionData.response_text || 
        "Thank you for bringing this to our attention. Our team is investigating this issue and will provide updates soon.";
      
      if (!this.config.enable_real_posting) {
        // Demo mode - create success card without actual posting
        return {
          success: true,
          message: '🎯 Community response ready for posting!',
          card: this.createPostPreviewCard(responseText, actionContext),
          next_actions: ['view_on_github', 'edit_response'],
          metadata: {
            action_type: 'post_response',
            user_impact: 'high',
            requires_follow_up: false,
            debug_info: {
              demo_mode: true,
              response_length: responseText.length
            }
          }
        };
      }

      // Production mode - actually post to GitHub
      const postResult = await this.postToGitHub(responseText, actionContext);
      
      if (postResult.success) {
        return {
          success: true,
          message: '✅ Community response posted successfully!',
          card: this.createPostSuccessCard(postResult.url || this.config.github_demo_issue_url || '', responseText),
          next_actions: ['view_post', 'track_engagement'],
          metadata: {
            action_type: 'post_response',
            user_impact: 'high',
            requires_follow_up: false,
            debug_info: {
              post_url: postResult.url,
              response_length: responseText.length
            }
          }
        };
      } else {
        throw new Error(postResult.error || 'Failed to post to GitHub');
      }
      
    } catch (error: any) {
      return {
        success: false,
        message: `Failed to post response: ${error.message}`,
        card: this.createPostErrorCard(error.message, actionData.response_text),
        metadata: {
          action_type: 'post_response',
          user_impact: 'medium',
          requires_follow_up: true,
          debug_info: { error: error.message }
        }
      };
    }
  }

  /**
   * Handle internal escalation
   */
  private async handleEscalateInternal(
    context: TurnContext,
    actionData: any,
    actionContext: ActionContext
  ): Promise<Omit<ActionResult, 'execution_time_ms'>> {
    
    try {
      // In production, this would integrate with internal ticketing systems
      const escalationData = {
        issue_id: `ESC-${Date.now()}`,
        priority: actionData.priority || 'high',
        description: actionData.escalation_reason || 'Critical pain points detected in community feedback',
        affected_components: actionData.affected_components || [],
        reporter: actionContext.user_id,
        context: actionContext.original_query
      };

      // Simulate internal escalation process
      const escalationResult = await this.simulateInternalEscalation(escalationData);

      return {
        success: true,
        message: `🚨 Issue escalated successfully! Tracking ID: ${escalationData.issue_id}`,
        card: this.createEscalationCard(escalationData),
        next_actions: ['track_escalation', 'notify_team'],
        metadata: {
          action_type: 'escalate',
          user_impact: 'high',
          requires_follow_up: true,
          debug_info: {
            escalation_id: escalationData.issue_id,
            priority: escalationData.priority,
            components_count: escalationData.affected_components.length
          }
        }
      };
      
    } catch (error: any) {
      return {
        success: false,
        message: `Escalation failed: ${error.message}`,
        card: this.createErrorCard('escalate', error.message),
        metadata: {
          action_type: 'escalate',
          user_impact: 'high',
          requires_follow_up: true,
          debug_info: { error: error.message }
        }
      };
    }
  }

  /**
   * Handle creating tracking issue
   */
  private async handleCreateTrackingIssue(
    context: TurnContext,
    actionData: any,
    actionContext: ActionContext
  ): Promise<Omit<ActionResult, 'execution_time_ms'>> {
    
    try {
      const trackingData = {
        issue_id: `TRACK-${Date.now()}`,
        title: actionData.issue_title || `Community feedback analysis - ${new Date().toISOString().split('T')[0]}`,
        pain_points_count: actionData.pain_points_count || 0,
        affected_components: actionData.affected_components || [],
        created_by: actionContext.user_id,
        source_query: actionContext.original_query
      };

      // Simulate tracking issue creation
      const trackingResult = await this.simulateTrackingIssueCreation(trackingData);

      return {
        success: true,
        message: `📝 Tracking issue created! ID: ${trackingData.issue_id}`,
        card: this.createTrackingIssueCard(trackingData),
        next_actions: ['view_issue', 'add_details'],
        metadata: {
          action_type: 'track_issue',
          user_impact: 'medium',
          requires_follow_up: false,
          debug_info: {
            tracking_id: trackingData.issue_id,
            pain_points_count: trackingData.pain_points_count
          }
        }
      };
      
    } catch (error: any) {
      return {
        success: false,
        message: `Failed to create tracking issue: ${error.message}`,
        card: this.createErrorCard('track_issue', error.message),
        metadata: {
          action_type: 'track_issue',
          user_impact: 'medium',
          requires_follow_up: true,
          debug_info: { error: error.message }
        }
      };
    }
  }

  /**
   * Handle documentation update request
   */
  private async handleUpdateDocumentation(
    context: TurnContext,
    actionData: any,
    actionContext: ActionContext
  ): Promise<Omit<ActionResult, 'execution_time_ms'>> {
    
    try {
      const docUpdateData = {
        request_id: `DOC-${Date.now()}`,
        areas_to_update: actionData.areas_to_update || [],
        priority_level: actionData.priority_level || 'medium',
        requestor: actionContext.user_id,
        source_analysis: actionContext.original_query
      };

      // Simulate documentation update workflow
      const updateResult = await this.simulateDocumentationUpdate(docUpdateData);

      return {
        success: true,
        message: `📚 Documentation update requested! ID: ${docUpdateData.request_id}`,
        card: this.createDocUpdateCard(docUpdateData),
        next_actions: ['track_progress', 'review_draft'],
        metadata: {
          action_type: 'update_docs',
          user_impact: 'medium',
          requires_follow_up: false,
          debug_info: {
            update_id: docUpdateData.request_id,
            areas_count: docUpdateData.areas_to_update.length
          }
        }
      };
      
    } catch (error: any) {
      return {
        success: false,
        message: `Documentation update failed: ${error.message}`,
        card: this.createErrorCard('update_docs', error.message),
        metadata: {
          action_type: 'update_docs',
          user_impact: 'low',
          requires_follow_up: true,
          debug_info: { error: error.message }
        }
      };
    }
  }

  /**
   * Handle retry analysis
   */
  private async handleRetryAnalysis(
    context: TurnContext,
    actionData: any,
    actionContext: ActionContext
  ): Promise<Omit<ActionResult, 'execution_time_ms'>> {
    
    try {
      // Send progress card immediately
      const progressCardResult = this.progressCard.generateCard({
        stage: 'ingestion',
        message: 'Retrying analysis with refined parameters...',
        progress_percentage: 10,
        current_operation: 'Preparing retry attempt'
      });
      
      await context.sendActivity(MessageFactory.attachment(progressCardResult.card));

      // Re-run the analysis (this would integrate with your main bot flow)
      // For now, simulate a retry
      const retryResult = await this.simulateRetryAnalysis(actionData, actionContext);

      return {
        success: true,
        message: '🔄 Analysis retried successfully!',
        card: this.createRetrySuccessCard(retryResult),
        next_actions: ['view_results', 'compare_attempts'],
        metadata: {
          action_type: 'retry_analysis',
          user_impact: 'high',
          requires_follow_up: false,
          debug_info: {
            retry_attempt: actionData.retry_attempt || 1,
            original_query: actionContext.original_query
          }
        }
      };
      
    } catch (error: any) {
      return {
        success: false,
        message: `Analysis retry failed: ${error.message}`,
        card: this.createErrorCard('retry_analysis', error.message),
        metadata: {
          action_type: 'retry_analysis',
          user_impact: 'medium',
          requires_follow_up: true,
          debug_info: { error: error.message }
        }
      };
    }
  }

  /**
   * Handle contact support
   */
  private async handleContactSupport(
    context: TurnContext,
    actionData: any,
    actionContext: ActionContext
  ): Promise<Omit<ActionResult, 'execution_time_ms'>> {
    
    const supportData = {
      ticket_id: `SUP-${Date.now()}`,
      context: actionData.context || 'general',
      user_query: actionContext.original_query,
      user_id: actionContext.user_id,
      priority: 'normal'
    };

    return {
      success: true,
      message: `📞 Support ticket created! ID: ${supportData.ticket_id}`,
      card: this.createSupportCard(supportData),
      next_actions: ['check_status', 'add_details'],
      metadata: {
        action_type: 'contact_support',
        user_impact: 'medium',
        requires_follow_up: false,
        debug_info: {
          ticket_id: supportData.ticket_id,
          context: supportData.context
        }
      }
    };
  }

  /**
   * Post to GitHub demo issue
   */
  private async postToGitHub(responseText: string, actionContext: ActionContext): Promise<{
    success: boolean;
    url?: string;
    error?: string;
  }> {
    try {
      if (!this.config.github_demo_issue_url) {
        throw new Error('GitHub demo issue URL not configured');
      }

      const commentBody = `**Community Insights Bot Response**\n\n${responseText}\n\n---\n*Posted via Community Insights Bot | Query: "${actionContext.original_query}" | ${new Date().toISOString()}*`;

      // Use fetch to post to GitHub API
      const response = await fetch(this.config.github_demo_issue_url, {
        method: 'POST',
        headers: {
          'Authorization': `token ${process.env.GITHUB_TOKEN}`,
          'Accept': 'application/vnd.github.v3+json',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          body: commentBody
        })
      });

      if (!response.ok) {
        throw new Error(`GitHub API returned ${response.status}: ${response.statusText}`);
      }

      const result = await response.json();
      
      return {
        success: true,
        url: result.html_url || this.config.github_demo_issue_url
      };
      
    } catch (error: any) {
      console.error('GitHub posting error:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  // Card creation methods
  private createPostPreviewCard(responseText: string, actionContext: ActionContext) {
    return {
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: '🎯 **Community Response Ready**',
          weight: 'bolder',
          size: 'medium'
        },
        {
          type: 'TextBlock',
          text: 'Preview of response that will be posted:',
          size: 'small'
        },
        {
          type: 'Container',
          style: 'emphasis',
          items: [{
            type: 'TextBlock',
            text: this.truncateText(responseText, 400),
            wrap: true,
            size: 'small'
          }]
        },
        {
          type: 'TextBlock',
          text: `Target: Demo GitHub Issue #4 | Query: "${actionContext.original_query}"`,
          size: 'small',
          color: 'accent'
        }
      ],
      actions: [
        {
          type: 'Action.OpenUrl',
          title: '👀 View Demo Issue',
          url: 'https://github.com/s-sridharan/devx-bot/issues/4'
        },
        {
          type: 'Action.Submit',
          title: '✏️ Edit Response',
          data: { action: 'edit_response', original_text: responseText }
        }
      ]
    };
  }

  private createPostSuccessCard(postUrl: string, responseText: string) {
    return {
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: '✅ **Response Posted Successfully!**',
          weight: 'bolder',
          size: 'medium',
          color: 'good'
        },
        {
          type: 'TextBlock',
          text: 'Your community response has been posted to GitHub.',
          wrap: true
        },
        {
          type: 'FactSet',
          facts: [
            { title: 'Posted to', value: 'GitHub Issue #4' },
            { title: 'Response length', value: `${responseText.length} characters` },
            { title: 'Posted at', value: new Date().toLocaleString() }
          ]
        }
      ],
      actions: [
        {
          type: 'Action.OpenUrl',
          title: '🔗 View Posted Response',
          url: postUrl
        }
      ]
    };
  }

  private createEscalationCard(escalationData: any) {
    return {
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: '🚨 **Issue Escalated**',
          weight: 'bolder',
          size: 'medium',
          color: 'attention'
        },
        {
          type: 'FactSet',
          facts: [
            { title: 'Escalation ID', value: escalationData.issue_id },
            { title: 'Priority', value: escalationData.priority },
            { title: 'Components', value: escalationData.affected_components.join(', ') || 'General' },
            { title: 'Status', value: 'Pending team assignment' }
          ]
        }
      ]
    };
  }

  private createTrackingIssueCard(trackingData: any) {
    return {
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: '📝 **Tracking Issue Created**',
          weight: 'bolder',
          size: 'medium'
        },
        {
          type: 'FactSet',
          facts: [
            { title: 'Issue ID', value: trackingData.issue_id },
            { title: 'Title', value: trackingData.title },
            { title: 'Pain Points', value: trackingData.pain_points_count.toString() },
            { title: 'Status', value: 'Open' }
          ]
        }
      ]
    };
  }

  private createDocUpdateCard(docUpdateData: any) {
    return {
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: '📚 **Documentation Update Requested**',
          weight: 'bolder',
          size: 'medium'
        },
        {
          type: 'FactSet',
          facts: [
            { title: 'Request ID', value: docUpdateData.request_id },
            { title: 'Priority', value: docUpdateData.priority_level },
            { title: 'Areas', value: docUpdateData.areas_to_update.join(', ') || 'Multiple areas' },
            { title: 'Status', value: 'Queued for review' }
          ]
        }
      ]
    };
  }

  private createRetrySuccessCard(retryResult: any) {
    return {
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: '🔄 **Analysis Retried Successfully**',
          weight: 'bolder',
          size: 'medium',
          color: 'good'
        },
        {
          type: 'TextBlock',
          text: 'The analysis has been re-run with improved parameters.',
          wrap: true
        }
      ]
    };
  }

  private createSupportCard(supportData: any) {
    return {
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: '📞 **Support Ticket Created**',
          weight: 'bolder',
          size: 'medium'
        },
        {
          type: 'FactSet',
          facts: [
            { title: 'Ticket ID', value: supportData.ticket_id },
            { title: 'Priority', value: supportData.priority },
            { title: 'Context', value: supportData.context },
            { title: 'Status', value: 'Open' }
          ]
        },
        {
          type: 'TextBlock',
          text: 'Our support team will respond within 24 hours.',
          size: 'small',
          color: 'accent'
        }
      ]
    };
  }

  private createErrorCard(actionType: string, errorMessage: string) {
    return {
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: '⚠️ **Action Failed**',
          weight: 'bolder',
          size: 'medium',
          color: 'attention'
        },
        {
          type: 'TextBlock',
          text: `Failed to execute ${actionType}: ${errorMessage}`,
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: '🔄 Retry',
          data: { action: actionType, retry: true }
        }
      ]
    };
  }

  private createPostErrorCard(errorMessage: string, responseText: string) {
    return {
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: '❌ **Posting Failed**',
          weight: 'bolder',
          size: 'medium',
          color: 'attention'
        },
        {
          type: 'TextBlock',
          text: `Could not post to GitHub: ${errorMessage}`,
          wrap: true
        },
        {
          type: 'TextBlock',
          text: 'You can copy the response below and post it manually:',
          size: 'small'
        },
        {
          type: 'Container',
          style: 'emphasis',
          items: [{
            type: 'TextBlock',
            text: responseText,
            wrap: true,
            size: 'small'
          }]
        }
      ],
      actions: [
        {
          type: 'Action.OpenUrl',
          title: '🔗 Open GitHub Issue',
          url: 'https://github.com/s-sridharan/devx-bot/issues/4'
        }
      ]
    };
  }

  // Simulation methods (replace with real integrations in production)
  private async simulateInternalEscalation(escalationData: any): Promise<any> {
    // Simulate API call delay
    await this.delay(500);
    return { success: true, assigned_team: 'Platform Engineering' };
  }

  private async simulateTrackingIssueCreation(trackingData: any): Promise<any> {
    await this.delay(300);
    return { success: true, issue_url: `https://internal-tracker.microsoft.com/issues/${trackingData.issue_id}` };
  }

  private async simulateDocumentationUpdate(docUpdateData: any): Promise<any> {
    await this.delay(400);
    return { success: true, workflow_id: `workflow-${Date.now()}` };
  }

  private async simulateRetryAnalysis(actionData: any, actionContext: ActionContext): Promise<any> {
    await this.delay(2000);
    return { 
      success: true, 
      pain_points_found: Math.floor(Math.random() * 5) + 1,
      confidence_improved: true 
    };
  }

  // Helper methods
  private validateActionData(actionData: any): void {
    if (!actionData || !actionData.action) {
      throw new Error('Invalid action data - missing action type');
    }

    const validActions = ['post_response', 'escalate', 'track_issue', 'update_docs', 'retry_analysis', 'contact_support'];
    if (!validActions.includes(actionData.action)) {
      throw new Error(`Invalid action type: ${actionData.action}`);
    }
  }

  private truncateText(text: string, maxLength: number): string {
    return text.length > maxLength ? `${text.substring(0, maxLength)}...` : text;
  }

  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Get current configuration
   */
  getConfig(): ActionHandlerConfig {
    return { ...this.config };
  }

  /**
   * Update configuration
   */
  updateConfig(newConfig: Partial<ActionHandlerConfig>): void {
    this.config = { ...this.config, ...newConfig };
  }
}