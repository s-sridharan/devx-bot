// src/agents/ResponseAgent.ts - Generate professional community responses from pain points
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */

import { IChatModel } from '@microsoft/teams.ai';
import { 
  ExtractedPainPoint,
  ResponseAgentInput, 
  ResponseAgentOutput, 
  AgentResult,
  CommunityResponse
} from '../types/index.js';
import { BugFixTimelines } from '../knowledge/BugFixTimelines.js';

export class ResponseAgent {
  private model: IChatModel;
  private responsePrompt: string;
  private knowledge: typeof BugFixTimelines;

  constructor(model: IChatModel) {
    this.model = model;
    this.knowledge = BugFixTimelines;
    this.responsePrompt = this.buildResponsePrompt();
  }

  /**
   * Execute response generation from extracted pain points
   */
  async execute(input: ResponseAgentInput): Promise<AgentResult<ResponseAgentOutput>> {
    const startTime = Date.now();
    const executionId = `response_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    
    try {
      console.log(`🎯 ResponseAgent: Generating community response for ${input.pain_points.length} pain points`);
      
      // Validate input
      this.validateInput(input);
      
      // Get knowledge base context
      const knowledgeContext = this.getKnowledgeContext(input.pain_points);
      
      // Generate base response using knowledge base templates
      const templateResponse = this.knowledge.generateResponse(input.pain_points, input.context);
      
      // Enhance response with AI if needed
      const enhancedResponse = await this.enhanceWithAI(
        input.pain_points, 
        templateResponse, 
        input, 
        knowledgeContext
      );
      
      // Generate suggested actions
      const suggestedActions = this.generateActions(input.pain_points, input.target_context);
      
      // Calculate response metadata
      const responseMetadata = this.calculateResponseMetadata(
        input.pain_points, 
        enhancedResponse, 
        knowledgeContext
      );
      
      const executionTime = Date.now() - startTime;
      
      console.log(`✅ ResponseAgent: Generated response with confidence ${enhancedResponse.confidence_level} (${executionTime}ms)`);
      
      return {
        success: true,
        data: {
          community_response: enhancedResponse,
          response_metadata: responseMetadata,
          suggested_actions: suggestedActions
        },
        execution_time_ms: executionTime,
        agent_name: 'ResponseAgent',
        metadata: {
          model_calls: responseMetadata.ai_enhancement_used ? 1 : 0,
          tokens_used: responseMetadata.ai_enhancement_used ? this.estimateTokens(enhancedResponse.response_text) : 0,
          confidence: enhancedResponse.confidence_level,
          debug_info: {
            execution_id: executionId,
            knowledge_base_matches: responseMetadata.knowledge_base_matches,
            template_used: responseMetadata.template_used,
            pain_points_processed: input.pain_points.length,
            actions_generated: suggestedActions.length
          }
        }
      };
      
    } catch (error: any) {
      const executionTime = Date.now() - startTime;
      console.error('❌ ResponseAgent Error:', error);
      
      return {
        success: false,
        error: `Response generation failed: ${error.message}`,
        execution_time_ms: executionTime,
        agent_name: 'ResponseAgent',
        metadata: {
          model_calls: 0,
          confidence: 0,
          debug_info: {
            execution_id: executionId,
            error_type: error.name || 'UnknownError',
            error_location: 'ResponseAgent.execute'
          }
        }
      };
    }
  }

  /**
   * Get knowledge base context for pain points
   */
  private getKnowledgeContext(painPoints: ExtractedPainPoint[]) {
    const categories = painPoints.map(pp => pp.category);
    const primaryCategory = this.getMostFrequent(categories) || 'engineering';
    
    // Get timeline data and team information
    const timelineData = this.knowledge.getTimelineData(primaryCategory);
    const responsibleTeam = this.knowledge.getResponsibleTeam(primaryCategory);
    const resolutionEstimate = this.knowledge.getEstimatedResolution(painPoints);
    
    return {
      primary_category: primaryCategory,
      timeline_data: timelineData,
      responsible_team: responsibleTeam,
      resolution_estimate: resolutionEstimate,
      available_categories: this.knowledge.getAvailableCategories()
    };
  }

  /**
   * Enhance template response with AI if needed
   */
  private async enhanceWithAI(
    painPoints: ExtractedPainPoint[],
    templateResponse: string,
    input: ResponseAgentInput,
    knowledgeContext: any
  ): Promise<CommunityResponse> {
    
    // Check if AI enhancement is needed
    const needsEnhancement = this.shouldEnhanceWithAI(input, templateResponse);
    
    if (!needsEnhancement) {
      return this.createResponseFromTemplate(templateResponse, painPoints);
    }

    try {
      // Build AI enhancement prompt
      const enhancementPrompt = this.buildEnhancementPrompt(
        painPoints, 
        templateResponse, 
        input, 
        knowledgeContext
      );
      
      // Get AI enhancement
      const result = await this.model.send({
        role: 'user',
        content: enhancementPrompt
      });
      
      const resultContent = result.content || templateResponse;
      
      // Parse and validate AI response
      return this.parseAIResponse(resultContent, painPoints, templateResponse);
      
    } catch (error) {
      console.warn('⚠️ AI enhancement failed, falling back to template response');
      return this.createResponseFromTemplate(templateResponse, painPoints);
    }
  }

  /**
   * Create community response from template
   */
  private createResponseFromTemplate(templateResponse: string, painPoints: ExtractedPainPoint[]): CommunityResponse {
    return {
      response_text: templateResponse,
      tone: 'professional',
      includes_workaround: templateResponse.includes('workaround') || templateResponse.includes('solution'),
      includes_timeline: templateResponse.includes('sprint') || templateResponse.includes('days'),
      tracking_references: this.extractTrackingReferences(templateResponse),
      confidence_level: 0.85, // High confidence for knowledge base templates
      estimated_helpfulness: this.estimateHelpfulness(templateResponse, painPoints)
    };
  }

  /**
   * Parse AI enhanced response
   */
  private parseAIResponse(content: string, painPoints: ExtractedPainPoint[], fallbackTemplate: string): CommunityResponse {
    try {
      // Try to parse as JSON first (if AI returns structured response)
      const cleanContent = content.replace(/```json\n?|\n?```/g, '').trim();
      
      let responseData: any;
      try {
        responseData = JSON.parse(cleanContent);
      } catch {
        // If not JSON, treat as plain text response
        responseData = { response_text: content };
      }
      
      return {
        response_text: responseData.response_text || content,
        tone: this.detectTone(responseData.response_text || content),
        includes_workaround: this.hasWorkaround(responseData.response_text || content),
        includes_timeline: this.hasTimeline(responseData.response_text || content),
        tracking_references: this.extractTrackingReferences(responseData.response_text || content),
        confidence_level: Math.max(0.1, Math.min(1.0, responseData.confidence_level || 0.7)),
        estimated_helpfulness: this.estimateHelpfulness(responseData.response_text || content, painPoints)
      };
      
    } catch (error) {
      console.warn('Failed to parse AI response, using fallback template');
      return this.createResponseFromTemplate(fallbackTemplate, painPoints);
    }
  }

  /**
   * Generate suggested actions based on pain points and context
   */
  private generateActions(painPoints: ExtractedPainPoint[], targetContext: any) {
    const actions = [];
    
    // Always offer to post response
    actions.push({
      action_type: 'post_response' as const,
      action_label: 'Post Community Response',
      action_data: {
        platform: targetContext.platform || 'github',
        ready_to_post: true
      },
      priority: 'high' as const
    });
    
    // Check for escalation needs
    const criticalPainPoints = painPoints.filter(pp => pp.priority === 'escalation_risk');
    if (criticalPainPoints.length > 0) {
      actions.push({
        action_type: 'escalate_internal' as const,
        action_label: `Escalate ${criticalPainPoints.length} Critical Issues`,
        action_data: {
          escalation_reason: 'Critical priority pain points detected',
          affected_components: [...new Set(criticalPainPoints.flatMap(pp => pp.affected_components))]
        },
        priority: 'high' as const
      });
    }
    
    // Check for documentation gaps
    const docPainPoints = painPoints.filter(pp => pp.category === 'documentation');
    if (docPainPoints.length > 2) {
      actions.push({
        action_type: 'update_documentation' as const,
        action_label: 'Update Documentation',
        action_data: {
          areas_to_update: docPainPoints.map(pp => pp.primary_pain_point),
          priority_level: 'medium'
        },
        priority: 'medium' as const
      });
    }
    
    // Check for tracking needs
    const engineeringPainPoints = painPoints.filter(pp => pp.category === 'engineering');
    if (engineeringPainPoints.length > 0) {
      actions.push({
        action_type: 'create_tracking_issue' as const,
        action_label: 'Create Internal Tracking Issue',
        action_data: {
          issue_title: `Engineering pain points analysis - ${new Date().toISOString().split('T')[0]}`,
          pain_points_count: engineeringPainPoints.length,
          affected_components: [...new Set(engineeringPainPoints.flatMap(pp => pp.affected_components))]
        },
        priority: 'medium' as const
      });
    }
    
    return actions;
  }

  /**
   * Calculate response metadata
   */
  private calculateResponseMetadata(
    painPoints: ExtractedPainPoint[], 
    response: CommunityResponse, 
    knowledgeContext: any
  ) {
    return {
      knowledge_base_matches: knowledgeContext.timeline_data ? [knowledgeContext.primary_category] : [],
      template_used: knowledgeContext.timeline_data?.subcategory || 'generic',
      personalization_applied: response.includes_timeline && response.includes_workaround,
      estimated_user_satisfaction: this.estimateUserSatisfaction(response, painPoints),
      ai_enhancement_used: response.confidence_level < 0.8 // Assume AI was used if confidence is lower
    };
  }

  /**
   * Build the core response generation prompt
   */
  private buildResponsePrompt(): string {
    return `You are a Microsoft Teams Platform community manager generating professional responses to developer feedback.

TASK: Enhance community response templates with specific technical details and personalization.

OUTPUT FORMAT: Return ONLY a JSON object with this structure:
{
  "response_text": "Professional community response with specific details",
  "confidence_level": 0.85,
  "tone": "professional",
  "includes_timeline": true,
  "includes_workaround": true
}

RESPONSE GUIDELINES:

Tone & Style:
- Professional but friendly, like a knowledgeable Microsoft PM
- Use specific technical details and component names
- Include emojis sparingly (🔧, 📚, 🚀) for visual appeal
- Reference specific versions, APIs, and tracking numbers when possible

Content Requirements:
- Acknowledge the specific pain point clearly
- Provide concrete workarounds with code snippets when applicable
- Include realistic timelines (sprints, release versions)
- Add tracking references and internal issue numbers
- Link to relevant documentation and resources

Technical Accuracy:
- Use accurate Microsoft Teams API terminology
- Reference real SDK versions and component names
- Provide working code examples when relevant
- Mention specific team ownership (Bot Framework Team, Identity Team, etc.)

IMPORTANT: Return ONLY the JSON object. No additional text or markdown formatting.`;
  }

  /**
   * Build AI enhancement prompt
   */
  private buildEnhancementPrompt(
    painPoints: ExtractedPainPoint[],
    templateResponse: string,
    input: ResponseAgentInput,
    knowledgeContext: any
  ): string {
    const painPointsSummary = painPoints.map(pp => ({
      category: pp.category,
      pain_point: pp.primary_pain_point,
      components: pp.affected_components,
      sentiment: pp.sentiment,
      priority: pp.priority
    }));

    return `${this.responsePrompt}

CONTEXT:
Target Platform: ${input.target_context.platform}
Target Audience: ${input.target_context.audience}
Urgency Level: ${input.target_context.urgency_level}

PAIN POINTS TO ADDRESS:
${JSON.stringify(painPointsSummary, null, 2)}

KNOWLEDGE BASE TEMPLATE:
${templateResponse}

ENHANCEMENT REQUIREMENTS:
- Max length: ${input.response_requirements.max_length} characters
- Include technical details: ${input.response_requirements.include_technical_details}
- Include workarounds: ${input.response_requirements.include_workarounds}
- Include timelines: ${input.response_requirements.include_timelines}

Enhance the template response with specific technical details and personalization while maintaining the professional tone.`;
  }

  // Helper methods
  private shouldEnhanceWithAI(input: ResponseAgentInput, templateResponse: string): boolean {
    // Enhance if specific requirements are requested or template is generic
    return input.response_requirements.include_technical_details ||
           input.target_context.urgency_level === 'immediate' ||
           templateResponse.length < 200 ||
           input.target_context.audience === 'enterprise_customer';
  }

  private detectTone(text: string): 'professional' | 'friendly' | 'technical' {
    if (text.includes('```') || text.includes('API') || text.includes('SDK')) return 'technical';
    if (text.includes('😊') || text.includes('Thanks!') || text.includes('Great question')) return 'friendly';
    return 'professional';
  }

  private hasWorkaround(text: string): boolean {
    return text.toLowerCase().includes('workaround') || 
           text.toLowerCase().includes('solution') || 
           text.includes('```') ||
           text.toLowerCase().includes('try this');
  }

  private hasTimeline(text: string): boolean {
    return text.includes('sprint') || 
           text.includes('days') || 
           text.includes('release') ||
           text.includes('v4.') ||
           text.includes('week');
  }

  private extractTrackingReferences(text: string): string[] {
    const matches = text.match(/#[\w-]+|\b\w+-\d+\b|issue #\d+/g) || [];
    return matches.slice(0, 5); // Limit to 5 references
  }

  private estimateHelpfulness(text: string, painPoints: ExtractedPainPoint[]): number {
    let score = 0.5; // Base score
    
    if (this.hasWorkaround(text)) score += 0.2;
    if (this.hasTimeline(text)) score += 0.2;
    if (text.includes('```')) score += 0.1; // Code snippets
    if (text.length > 300) score += 0.1; // Detailed response
    
    // Adjust based on pain point severity
    const criticalCount = painPoints.filter(pp => pp.priority === 'escalation_risk').length;
    if (criticalCount > 0 && text.includes('escalat')) score += 0.1;
    
    return Math.min(1.0, score);
  }

  private estimateUserSatisfaction(response: CommunityResponse, painPoints: ExtractedPainPoint[]): number {
    let satisfaction = response.estimated_helpfulness;
    
    // Boost for addressing high-priority pain points
    const highPriorityCount = painPoints.filter(pp => 
      pp.priority === 'escalation_risk' || pp.priority === 'sentiment_spike'
    ).length;
    
    if (highPriorityCount > 0 && response.includes_timeline) {
      satisfaction += 0.1;
    }
    
    return Math.min(1.0, satisfaction);
  }

  private getMostFrequent<T>(arr: T[]): T | null {
    if (arr.length === 0) return null;
    const frequency = arr.reduce((acc, item) => {
      acc[item as string] = (acc[item as string] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);
    
    return Object.keys(frequency).reduce((a, b) => 
      frequency[a] > frequency[b] ? a : b
    ) as T;
  }

  private validateInput(input: ResponseAgentInput): void {
    if (!input.pain_points || input.pain_points.length === 0) {
      throw new Error('No pain points provided for response generation');
    }
    
    if (input.pain_points.length > 20) {
      throw new Error('Too many pain points - maximum 20 allowed per response');
    }
    
    if (!input.target_context || !input.response_requirements) {
      throw new Error('Missing required context or requirements for response generation');
    }
  }

  private estimateTokens(text: string): number {
    // Rough estimate: 1 token ≈ 4 characters
    return Math.ceil(text.length / 4);
  }
}