// src/cards/ProgressCard.ts - Generate Teams Adaptive Cards for agent chain progress tracking
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */

import { ProgressUpdate } from '../types/index.js';

export interface ProgressCardOptions {
  show_technical_details?: boolean;
  compact_mode?: boolean;
  theme?: 'default' | 'dark' | 'high-contrast';
  include_time_estimates?: boolean;
}

export interface ProgressCardResult {
  card: any; // Adaptive Card JSON
  should_replace_previous?: boolean;
  auto_dismiss_after_ms?: number;
  accessibility_text: string;
}

export class ProgressCard {
  private options: ProgressCardOptions;

  constructor(options: ProgressCardOptions = {}) {
    this.options = {
      show_technical_details: false,
      compact_mode: false,
      theme: 'default',
      include_time_estimates: true,
      ...options
    };
  }

  /**
   * Generate progress card for current stage
   */
  generateCard(update: ProgressUpdate, context?: {
    user_query?: string;
    items_count?: number;
    started_at?: string;
  }): ProgressCardResult {
    
    const stageInfo = this.getStageInfo(update.stage);
    const progressBarColor = this.getProgressColor(update.progress_percentage);
    const isComplete = update.stage === 'complete';
    
    // Build the main card structure
    const card = {
      type: 'AdaptiveCard',
      version: '1.4',
      schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      body: [
        // Header Section
        this.buildHeaderSection(update, stageInfo, context),
        
        // Progress Section
        this.buildProgressSection(update, progressBarColor),
        
        // Details Section (if enabled)
        ...(this.options.show_technical_details ? [this.buildDetailsSection(update, context)] : []),
        
        // Time Estimates Section
        ...(this.options.include_time_estimates && update.estimated_time_remaining_ms ? 
            [this.buildTimeSection(update)] : []),
        
        // Footer Section (for complete stage)
        ...(isComplete ? [this.buildCompleteSection()] : [])
      ],
      ...(this.options.theme !== 'default' ? { 
        style: this.options.theme === 'dark' ? 'emphasis' : 'default' 
      } : {})
    };

    return {
      card,
      should_replace_previous: true, // Always replace previous progress card
      auto_dismiss_after_ms: isComplete ? 5000 : undefined, // Auto-dismiss completion card
      accessibility_text: this.buildAccessibilityText(update, stageInfo)
    };
  }

  /**
   * Generate error card for failed operations
   */
  generateErrorCard(error: {
    stage: string;
    message: string;
    can_retry?: boolean;
    technical_details?: string;
  }): ProgressCardResult {
    
    const card = {
      type: 'AdaptiveCard',
      version: '1.2',
      schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      body: [
        // Error Header
        {
          type: 'ColumnSet',
          columns: [
            {
              type: 'Column',
              width: 'auto',
              items: [{
                type: 'Image',
                url: 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjQiIGhlaWdodD0iMjQiIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTEyIDJMMTMuMDkgOC4yNkwyMCA5TDEzLjA5IDE1Ljc0TDEyIDIyTDEwLjkxIDE1Ljc0TDQgOUwxMC45MSA4LjI2TDEyIDJaIiBmaWxsPSIjRkY0NDQyIi8+CjxwYXRoIGQ9Ik0xMiA4VjE2IiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIvPgo8cGF0aCBkPSJNMTIgMThIOC4wMDAwMSIgc3Ryb2tlPSJ3aGl0ZSIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiLz4KPC9zdmc+Cg==',
                width: '24px',
                height: '24px'
              }]
            },
            {
              type: 'Column',
              width: 'stretch',
              items: [
                {
                  type: 'TextBlock',
                  text: '⚠️ Analysis Error',
                  weight: 'bolder',
                  size: 'medium',
                  color: 'attention'
                },
                {
                  type: 'TextBlock',
                  text: error.message,
                  wrap: true,
                  size: 'small'
                }
              ]
            }
          ]
        },

        // Technical Details (if available)
        ...(error.technical_details ? [{
          type: 'Container',
          style: 'emphasis',
          items: [{
            type: 'TextBlock',
            text: `**Technical Details:**\n${error.technical_details}`,
            wrap: true,
            size: 'small',
            fontType: 'monospace'
          }]
        }] : []),

        // Action Buttons
        {
          type: 'ActionSet',
          actions: [
            ...(error.can_retry ? [{
              type: 'Action.Submit',
              title: '🔄 Retry Analysis',
              data: {
                action: 'retry_analysis',
                stage: error.stage
              },
              style: 'positive'
            }] : []),
            {
              type: 'Action.Submit',
              title: '📞 Contact Support',
              data: {
                action: 'contact_support',
                error_context: error.stage
              }
            }
          ]
        }
      ]
    };

    return {
      card,
      should_replace_previous: true,
      accessibility_text: `Analysis error occurred during ${error.stage}: ${error.message}`
    };
  }

  /**
   * Build header section with stage info and branding
   */
  private buildHeaderSection(update: ProgressUpdate, stageInfo: any, context?: any) {
    return {
      type: 'ColumnSet',
      columns: [
        {
          type: 'Column',
          width: 'auto',
          items: [{
            type: 'Image',
            url: stageInfo.icon_data_url,
            width: '32px',
            height: '32px'
          }]
        },
        {
          type: 'Column',
          width: 'stretch',
          items: [
            {
              type: 'TextBlock',
              text: stageInfo.title,
              weight: 'bolder',
              size: 'medium'
            },
            {
              type: 'TextBlock',
              text: update.message,
              wrap: true,
              size: 'small',
              color: 'default'
            },
            ...(context?.user_query ? [{
              type: 'TextBlock',
              text: `Query: "${context.user_query}"`,
              wrap: true,
              size: 'small',
              color: 'accent',
              fontType: 'monospace'
            }] : [])
          ]
        }
      ]
    };
  }

  /**
   * Build progress bar section
   */
  private buildProgressSection(update: ProgressUpdate, progressBarColor: string) {
    if (this.options.compact_mode) {
      const stageInfo = this.getStageInfo(update.stage);
      return {
        type: 'TextBlock',
        text: `${stageInfo.emoji} ${Math.round(update.progress_percentage)}% complete`,
        size: 'small'
      };
    }

    // Calculate progress bar segments (20 segments for 5% increments)
    const segments = 20;
    const filledSegments = Math.round((update.progress_percentage / 100) * segments);
    
    return {
      type: 'Container',
      items: [
        {
          type: 'ColumnSet',
          columns: [
            {
              type: 'Column',
              width: 'stretch',
              items: [{
                type: 'TextBlock',
                text: this.buildProgressBar(filledSegments, segments, progressBarColor),
                fontType: 'monospace',
                size: 'small'
              }]
            },
            {
              type: 'Column',
              width: 'auto',
              items: [{
                type: 'TextBlock',
                text: `${Math.round(update.progress_percentage)}%`,
                weight: 'bolder',
                size: 'small'
              }]
            }
          ]
        }
      ]
    };
  }

  /**
   * Build technical details section
   */
  private buildDetailsSection(update: ProgressUpdate, context?: any) {
    const details = [];
    
    if (update.current_operation) {
      details.push(`**Current Operation:** ${update.current_operation}`);
    }
    
    if (context?.items_count) {
      details.push(`**Items to Process:** ${context.items_count}`);
    }
    
    if (context?.started_at) {
      const elapsed = Math.round((Date.now() - new Date(context.started_at).getTime()) / 1000);
      details.push(`**Elapsed Time:** ${elapsed}s`);
    }

    if (details.length === 0) return null;

    return {
      type: 'Container',
      style: 'emphasis',
      items: [{
        type: 'TextBlock',
        text: details.join('\n'),
        wrap: true,
        size: 'small'
      }]
    };
  }

  /**
   * Build time estimates section
   */
  private buildTimeSection(update: ProgressUpdate) {
    const remainingSeconds = Math.ceil((update.estimated_time_remaining_ms || 0) / 1000);
    const timeText = remainingSeconds > 60 
      ? `~${Math.ceil(remainingSeconds / 60)} minutes remaining`
      : `~${remainingSeconds} seconds remaining`;

    return {
      type: 'Container',
      items: [{
        type: 'TextBlock',
        text: `⏱️ ${timeText}`,
        size: 'small',
        color: 'accent',
        horizontalAlignment: 'center'
      }]
    };
  }

  /**
   * Build completion section with next steps
   */
  private buildCompleteSection() {
    return {
      type: 'Container',
      style: 'good',
      items: [
        {
          type: 'TextBlock',
          text: '🎉 **Analysis Complete!**',
          weight: 'bolder',
          size: 'medium',
          horizontalAlignment: 'center'
        },
        {
          type: 'TextBlock',
          text: 'Review the insights below and choose your next action.',
          wrap: true,
          size: 'small',
          horizontalAlignment: 'center'
        }
      ]
    };
  }

  /**
   * Get stage-specific information
   */
  private getStageInfo(stage: string) {
    const stageMap = {
      ingestion: {
        title: '🔍 Ingesting Feedback',
        emoji: '🔍',
        icon_data_url: 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMzIiIGhlaWdodD0iMzIiIHZpZXdCb3g9IjAgMCAzMiAzMiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGNpcmNsZSBjeD0iMTYiIGN5PSIxNiIgcj0iMTQiIGZpbGw9IiM0Qzg0RkYiLz4KPHBhdGggZD0iTTEyIDEySDIwVjIwSDEyVjEyWiIgc3Ryb2tlPSJ3aGl0ZSIgc3Ryb2tlLXdpZHRoPSIyIi8+CjxwYXRoIGQ9Ik0xNiAxNkwyMCAyMCIgc3Ryb2tlPSJ3aGl0ZSIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiLz4KPC9zdmc+Cg=='
      },
      extraction: {
        title: '🧠 Analyzing Pain Points',
        emoji: '🧠',
        icon_data_url: 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMzIiIGhlaWdodD0iMzIiIHZpZXdCb3g9IjAgMCAzMiAzMiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGNpcmNsZSBjeD0iMTYiIGN5PSIxNiIgcj0iMTQiIGZpbGw9IiNGRjhDMDAiLz4KPHBhdGggZD0iTTEwIDEySDIyVjE0SDEwVjEyWiIgZmlsbD0id2hpdGUiLz4KPHBhdGggZD0iTTEwIDE2SDIyVjE4SDEwVjE2WiIgZmlsbD0id2hpdGUiLz4KPHBhdGggZD0iTTEwIDIwSDIyVjIySDEwVjIwWiIgZmlsbD0id2hpdGUiLz4KPC9zdmc+Cg=='
      },
      response_generation: {
        title: '✍️ Generating Response',
        emoji: '✍️',
        icon_data_url: 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMzIiIGhlaWdodD0iMzIiIHZpZXdCb3g9IjAgMCAzMiAzMiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGNpcmNsZSBjeD0iMTYiIGN5PSIxNiIgcj0iMTQiIGZpbGw9IiMyNUQzNjYiLz4KPHBhdGggZD0iTTEwIDEwSDIyVjIySDEwVjEwWiIgc3Ryb2tlPSJ3aGl0ZSIgc3Ryb2tlLXdpZHRoPSIyIiBmaWxsPSJub25lIi8+CjxwYXRoIGQ9Ik0xMyAxM0gxOSIgc3Ryb2tlPSJ3aGl0ZSIgc3Ryb2tlLXdpZHRoPSIxLjUiLz4KPHBhdGggZD0iTTEzIDE2SDE5IiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjEuNSIvPgo8cGF0aCBkPSJNMTMgMTlIMTciIHN0cm9rZT0id2hpdGUiIHN0cm9rZS13aWR0aD0iMS41Ii8+Cjwvc3ZnPgo='
      },
      complete: {
        title: '✅ Analysis Complete',
        emoji: '✅',
        icon_data_url: 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMzIiIGhlaWdodD0iMzIiIHZpZXdCb3g9IjAgMCAzMiAzMiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGNpcmNsZSBjeD0iMTYiIGN5PSIxNiIgcj0iMTQiIGZpbGw9IiMxNkExNzUiLz4KPHBhdGggZD0iTTEwIDE2TDE0IDIwTDIyIDEyIiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjMiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCIvPgo8L3N2Zz4K'
      }
    };

    return stageMap[stage as keyof typeof stageMap] || stageMap.ingestion;
  }

  /**
   * Get progress bar color based on percentage
   */
  private getProgressColor(percentage: number): string {
    if (percentage >= 100) return '#16A175'; // Success green
    if (percentage >= 70) return '#25D366';  // Progress green
    if (percentage >= 30) return '#FF8C00';  // Warning orange
    return '#4C84FF';  // Info blue
  }

  /**
   * Build ASCII progress bar
   */
  private buildProgressBar(filled: number, total: number, color: string): string {
    const filledChar = '█';
    const emptyChar = '░';
    
    const filledPart = filledChar.repeat(filled);
    const emptyPart = emptyChar.repeat(total - filled);
    
    return `${filledPart}${emptyPart}`;
  }

  /**
   * Build accessibility text for screen readers
   */
  private buildAccessibilityText(update: ProgressUpdate, stageInfo: any): string {
    const stage = stageInfo.title.replace(/[^\w\s]/g, ''); // Remove emojis
    const progress = Math.round(update.progress_percentage);
    
    let text = `${stage}. ${progress} percent complete. ${update.message}`;
    
    if (update.current_operation) {
      text += ` Currently ${update.current_operation.toLowerCase()}.`;
    }
    
    if (update.estimated_time_remaining_ms) {
      const remainingSeconds = Math.ceil(update.estimated_time_remaining_ms / 1000);
      text += ` Estimated ${remainingSeconds} seconds remaining.`;
    }
    
    return text;
  }

  /**
   * Generate compact progress card for mobile/small screens
   */
  generateCompactCard(update: ProgressUpdate): ProgressCardResult {
    const originalCompactMode = this.options.compact_mode;
    this.options.compact_mode = true;
    
    const result = this.generateCard(update);
    
    this.options.compact_mode = originalCompactMode;
    return result;
  }

  /**
   * Generate progress card specifically for Teams mobile
   */
  generateMobileCard(update: ProgressUpdate): ProgressCardResult {
    const stageInfo = this.getStageInfo(update.stage);
    
    const card = {
      type: 'AdaptiveCard',
      version: '1.4',
      schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      body: [
        {
          type: 'TextBlock',
          text: `${stageInfo.emoji} ${stageInfo.title}`,
          weight: 'bolder',
          size: 'medium'
        },
        {
          type: 'TextBlock',
          text: `${Math.round(update.progress_percentage)}% - ${update.message}`,
          wrap: true,
          size: 'small'
        }
      ]
    };

    return {
      card,
      should_replace_previous: true,
      accessibility_text: this.buildAccessibilityText(update, stageInfo)
    };
  }

  /**
   * Update options
   */
  updateOptions(newOptions: Partial<ProgressCardOptions>): void {
    this.options = { ...this.options, ...newOptions };
  }

  /**
   * Get current options
   */
  getOptions(): ProgressCardOptions {
    return { ...this.options };
  }
}