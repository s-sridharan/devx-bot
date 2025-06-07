// src/orchestration/AgentChain.ts - Orchestrate extraction and response agent pipeline
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */

import { IChatModel } from '@microsoft/teams.ai';
import { 
  RawFeedbackItem,
  AgentChainConfig,
  AgentChainResult,
  AgentExecutionContext,
  ProgressUpdate,
  ProgressCallback,
  ExtractionAgentInput,
  ResponseAgentInput,
  ExtractionAgentOutput,
  ResponseAgentOutput,
  ProcessingError
} from '../types/index.js';
import { ExtractionAgent } from '../agents/ExtractionAgent.js';
import { ResponseAgent } from '../agents/ResponseAgent.js';

export interface AgentChainInput {
  raw_feedback: RawFeedbackItem[];
  target_context: {
    platform: 'github' | 'stackoverflow' | 'general';
    audience: 'developer_community' | 'enterprise_customer' | 'internal_team';
    urgency_level: 'immediate' | 'normal' | 'low_priority';
  };
  response_requirements: {
    include_technical_details: boolean;
    include_workarounds: boolean;
    include_timelines: boolean;
    max_length: number;
  };
  execution_context: AgentExecutionContext;
  progress_callback?: ProgressCallback;
}

export class AgentChain {
  private extractionAgent: ExtractionAgent;
  private responseAgent: ResponseAgent;
  private config: AgentChainConfig;
  private model: IChatModel;

  constructor(model: IChatModel, config?: Partial<AgentChainConfig>) {
    this.model = model;
    this.config = this.buildDefaultConfig(config);
    this.extractionAgent = new ExtractionAgent(model);
    this.responseAgent = new ResponseAgent(model);
  }

  /**
   * Execute the complete agent chain: Extraction → Response Generation
   */
  async execute(input: AgentChainInput): Promise<AgentChainResult> {
    const startTime = Date.now();
    const chainId = `chain_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    
    let stepsCompleted = 0;
    let stepsFailed = 0;
    let extractionResult: any = null;
    let responseResult: any = null;

    try {
      console.log(`🚀 AgentChain: Starting execution for ${input.raw_feedback.length} feedback items`);
      
      // Validate input
      this.validateInput(input);
      
      // Stage 1: Ingestion Complete (already done by tools)
      await this.updateProgress(input.progress_callback, {
        stage: 'ingestion',
        message: `Processing ${input.raw_feedback.length} feedback items`,
        progress_percentage: 10,
        current_operation: 'Data validation and preparation'
      });

      // Stage 2: Pain Point Extraction
      await this.updateProgress(input.progress_callback, {
        stage: 'extraction',
        message: 'Analyzing feedback for pain points and sentiment',
        progress_percentage: 30,
        estimated_time_remaining_ms: 15000,
        current_operation: 'Running AI extraction analysis'
      });

      extractionResult = await this.executeExtractionWithRetry(input);
      
      if (extractionResult.success) {
        stepsCompleted++;
        console.log(`✅ Extraction completed: ${extractionResult.data.pain_points.length} pain points found`);
      } else {
        stepsFailed++;
        console.error('❌ Extraction failed:', extractionResult.error);
      }

      // Stage 3: Response Generation
      await this.updateProgress(input.progress_callback, {
        stage: 'response_generation',
        message: 'Generating professional community response',
        progress_percentage: 70,
        estimated_time_remaining_ms: 8000,
        current_operation: 'Applying knowledge base templates and AI enhancement'
      });

      if (extractionResult.success && extractionResult.data.pain_points.length > 0) {
        responseResult = await this.executeResponseWithRetry(input, extractionResult.data);
        
        if (responseResult.success) {
          stepsCompleted++;
          console.log(`✅ Response generated with confidence ${responseResult.data.community_response.confidence_level}`);
        } else {
          stepsFailed++;
          console.error('❌ Response generation failed:', responseResult.error);
        }
      } else {
        // Create fallback response result
        responseResult = this.createFallbackResponse(input.execution_context);
        stepsFailed++;
      }

      // Stage 4: Complete
      await this.updateProgress(input.progress_callback, {
        stage: 'complete',
        message: 'Analysis complete - ready for action',
        progress_percentage: 100,
        current_operation: 'Finalizing results and generating action suggestions'
      });

      const totalExecutionTime = Date.now() - startTime;
      const chainMetadata = this.calculateChainMetadata(
        extractionResult, 
        responseResult, 
        stepsCompleted, 
        stepsFailed,
        totalExecutionTime
      );

      console.log(`🎯 AgentChain: Completed in ${totalExecutionTime}ms - ${stepsCompleted}/2 steps successful`);

      return {
        extraction_result: extractionResult,
        response_result: responseResult,
        overall_success: stepsCompleted >= 1, // Success if at least extraction worked
        total_execution_time_ms: totalExecutionTime,
        chain_metadata: chainMetadata
      };

    } catch (error: any) {
      const totalExecutionTime = Date.now() - startTime;
      console.error('❌ AgentChain Fatal Error:', error);

      // Create error response
      const errorResult = this.createErrorResponse(error, input.execution_context, totalExecutionTime);
      
      return {
        extraction_result: extractionResult || errorResult,
        response_result: responseResult || errorResult,
        overall_success: false,
        total_execution_time_ms: totalExecutionTime,
        chain_metadata: {
          steps_completed: stepsCompleted,
          steps_failed: stepsFailed + 1,
          quality_score: 0,
          user_experience_score: 0
        }
      };
    }
  }

  /**
   * Execute extraction agent with retry logic
   */
  private async executeExtractionWithRetry(input: AgentChainInput) {
    const maxRetries = this.config.error_handling.max_retries;
    let lastError: any = null;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        const extractionInput: ExtractionAgentInput = {
          raw_feedback: input.raw_feedback,
          analysis_focus: {
            extract_technical_details: input.response_requirements.include_technical_details,
            prioritize_recent: true,
            group_similar_issues: true
          },
          context: input.execution_context
        };

        const result = await this.extractionAgent.execute(extractionInput);
        
        if (result.success) {
          return result;
        } else {
          lastError = new Error(result.error);
          console.warn(`⚠️ Extraction attempt ${attempt}/${maxRetries} failed: ${result.error}`);
        }

      } catch (error: any) {
        lastError = error;
        console.warn(`⚠️ Extraction attempt ${attempt}/${maxRetries} threw error: ${error.message}`);
      }

      // Wait before retry (exponential backoff)
      if (attempt < maxRetries) {
        const delay = Math.min(1000 * Math.pow(2, attempt - 1), 5000);
        await this.delay(delay);
      }
    }

    // All retries failed
    return {
      success: false,
      error: `Extraction failed after ${maxRetries} attempts: ${lastError?.message || 'Unknown error'}`,
      execution_time_ms: 0,
      agent_name: 'ExtractionAgent',
      metadata: {
        model_calls: 0,
        confidence: 0,
        debug_info: {
          retry_attempts: maxRetries,
          final_error: lastError?.message
        }
      }
    };
  }

  /**
   * Execute response agent with retry logic
   */
  private async executeResponseWithRetry(input: AgentChainInput, extractionData: ExtractionAgentOutput) {
    const maxRetries = this.config.error_handling.max_retries;
    let lastError: any = null;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        const responseInput: ResponseAgentInput = {
          pain_points: extractionData.pain_points,
          target_context: input.target_context,
          response_requirements: input.response_requirements,
          context: input.execution_context
        };

        const result = await this.responseAgent.execute(responseInput);
        
        if (result.success) {
          return result;
        } else {
          lastError = new Error(result.error);
          console.warn(`⚠️ Response attempt ${attempt}/${maxRetries} failed: ${result.error}`);
        }

      } catch (error: any) {
        lastError = error;
        console.warn(`⚠️ Response attempt ${attempt}/${maxRetries} threw error: ${error.message}`);
      }

      // Wait before retry
      if (attempt < maxRetries) {
        const delay = Math.min(1000 * Math.pow(2, attempt - 1), 5000);
        await this.delay(delay);
      }
    }

    // All retries failed - create fallback response
    return this.createFallbackResponse(input.execution_context, extractionData.pain_points);
  }

  /**
   * Create fallback response when response generation fails
   */
  private createFallbackResponse(context: AgentExecutionContext, painPoints?: any[]) {
    const fallbackText = painPoints && painPoints.length > 0
      ? `Thank you for bringing these ${painPoints.length} issues to our attention. Our team is reviewing the feedback and will provide updates soon.`
      : this.config.error_handling.fallback_responses.general || 
        'Thank you for your feedback. Our team is investigating these issues and will provide updates as more information becomes available.';

    return {
      success: true,
      data: {
        community_response: {
          response_text: fallbackText,
          tone: 'professional' as const,
          includes_workaround: false,
          includes_timeline: false,
          tracking_references: [],
          confidence_level: 0.3,
          estimated_helpfulness: 0.4
        },
        response_metadata: {
          knowledge_base_matches: [],
          template_used: 'fallback',
          personalization_applied: false,
          estimated_user_satisfaction: 0.3,
          ai_enhancement_used: false
        },
        suggested_actions: [{
          action_type: 'escalate_internal' as const,
          action_label: 'Manual Review Required',
          action_data: { reason: 'Automated response generation failed' },
          priority: 'high' as const
        }]
      },
      execution_time_ms: 0,
      agent_name: 'ResponseAgent',
      metadata: {
        model_calls: 0,
        confidence: 0.3,
        debug_info: {
          fallback_used: true,
          context_id: context.request_id
        }
      }
    };
  }

  /**
   * Create error response for fatal errors
   */
  private createErrorResponse(error: any, context: AgentExecutionContext, executionTime: number) {
    return {
      success: false,
      error: `Chain execution failed: ${error.message}`,
      execution_time_ms: executionTime,
      agent_name: 'AgentChain',
      metadata: {
        model_calls: 0,
        confidence: 0,
        debug_info: {
          error_type: error.name || 'UnknownError',
          context_id: context.request_id,
          fatal_error: true
        }
      }
    };
  }

  /**
   * Calculate chain metadata and quality scores
   */
  private calculateChainMetadata(
    extractionResult: any, 
    responseResult: any, 
    stepsCompleted: number, 
    stepsFailed: number,
    totalExecutionTime: number
  ) {
    // Calculate quality score based on extraction confidence and response quality
    let qualityScore = 0;
    if (extractionResult?.success) {
      qualityScore += extractionResult.metadata.confidence * 0.5;
    }
    if (responseResult?.success) {
      qualityScore += responseResult.data.community_response.confidence_level * 0.5;
    }

    // Calculate user experience score based on execution time and completeness
    let userExperienceScore = Math.max(0, 1 - (totalExecutionTime / 30000)); // Penalty for >30s
    if (stepsCompleted === 2) userExperienceScore += 0.3; // Bonus for complete execution
    if (stepsFailed === 0) userExperienceScore += 0.2; // Bonus for no failures

    return {
      steps_completed: stepsCompleted,
      steps_failed: stepsFailed,
      quality_score: Math.min(1.0, qualityScore),
      user_experience_score: Math.min(1.0, userExperienceScore)
    };
  }

  /**
   * Update progress via callback
   */
  private async updateProgress(callback: ProgressCallback | undefined, update: ProgressUpdate): Promise<void> {
    if (callback) {
      try {
        await callback(update);
      } catch (error) {
        console.warn('Progress callback failed:', error);
      }
    }
  }

  /**
   * Build default configuration
   */
  private buildDefaultConfig(config?: Partial<AgentChainConfig>): AgentChainConfig {
    return {
      extraction_agent: {
        model_config: {
          temperature: 0.1,
          max_tokens: 2000,
          timeout_ms: 15000
        },
        prompt_template: 'default_extraction',
        validation_rules: ['minimum_confidence_0.5', 'require_category', 'require_sentiment']
      },
      response_agent: {
        model_config: {
          temperature: 0.3,
          max_tokens: 1000,
          timeout_ms: 10000
        },
        knowledge_base_enabled: true,
        response_templates_enabled: true
      },
      error_handling: {
        max_retries: 2,
        fallback_responses: {
          general: 'Thank you for your feedback. Our team is investigating and will provide updates soon.',
          extraction_failed: 'We\'re having trouble analyzing the feedback data. Please try again or contact support.',
          response_failed: 'Thank you for the feedback. Our team will review manually and respond appropriately.'
        },
        escalation_threshold: 3
      },
      ...config
    };
  }

  /**
   * Validate chain input
   */
  private validateInput(input: AgentChainInput): void {
    if (!input.raw_feedback || input.raw_feedback.length === 0) {
      throw new Error('No feedback data provided for agent chain execution');
    }

    if (input.raw_feedback.length > 100) {
      throw new Error('Too much feedback data - maximum 100 items allowed per chain execution');
    }

    if (!input.execution_context?.request_id) {
      throw new Error('Missing execution context or request ID');
    }

    if (!input.target_context || !input.response_requirements) {
      throw new Error('Missing target context or response requirements');
    }
  }

  /**
   * Simple delay utility
   */
  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Get current configuration
   */
  getConfig(): AgentChainConfig {
    return { ...this.config };
  }

  /**
   * Update configuration
   */
  updateConfig(newConfig: Partial<AgentChainConfig>): void {
    this.config = { ...this.config, ...newConfig };
  }

  /**
   * Health check for the entire chain
   */
  async healthCheck(): Promise<{
    healthy: boolean;
    message: string;
    component_status: {
      extraction_agent: boolean;
      response_agent: boolean;
      model_connectivity: boolean;
    };
  }> {
    try {
      // Test basic model connectivity
      const testResult = await this.model.send({
        role: 'user',
        content: 'Test connectivity - respond with "OK"'
      });

      const modelHealthy = testResult.content?.includes('OK') || testResult.content !== undefined;

      return {
        healthy: modelHealthy,
        message: modelHealthy 
          ? 'Agent chain is healthy and ready for execution'
          : 'Model connectivity issues detected',
        component_status: {
          extraction_agent: true, // Agents are healthy if constructed successfully
          response_agent: true,
          model_connectivity: modelHealthy
        }
      };

    } catch (error: any) {
      return {
        healthy: false,
        message: `Health check failed: ${error.message}`,
        component_status: {
          extraction_agent: true,
          response_agent: true,
          model_connectivity: false
        }
      };
    }
  }
}