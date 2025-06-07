// src/agents/ExtractionAgent.ts - Transform raw feedback into structured pain points
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */

import { IChatModel } from '@microsoft/teams.ai';
import { 
  RawFeedbackItem, 
  ExtractedPainPoint, 
  ExtractionAgentInput, 
  ExtractionAgentOutput, 
  AgentResult,
  PerformanceMetrics 
} from '../types/index.js';

export class ExtractionAgent {
  private model: IChatModel;
  private extractionPrompt: string;

  constructor(model: IChatModel) {
    this.model = model;
    this.extractionPrompt = this.buildExtractionPrompt();
  }

  /**
   * Execute extraction analysis on raw feedback data
   */
  async execute(input: ExtractionAgentInput): Promise<AgentResult<ExtractionAgentOutput>> {
    const startTime = Date.now();
    const executionId = `extraction_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    
    try {
      console.log(`🔍 ExtractionAgent: Starting analysis of ${input.raw_feedback.length} items`);
      
      // Validate input
      this.validateInput(input);
      
      // Process feedback items in batches for better performance
      const batchSize = 5; // Process 5 items at a time to avoid token limits
      const allPainPoints: ExtractedPainPoint[] = [];
      let modelCalls = 0;
      let totalTokensUsed = 0;
      
      for (let i = 0; i < input.raw_feedback.length; i += batchSize) {
        const batch = input.raw_feedback.slice(i, i + batchSize);
        console.log(`📊 Processing batch ${Math.floor(i / batchSize) + 1}/${Math.ceil(input.raw_feedback.length / batchSize)}`);
        
        const batchResults = await this.processBatch(batch, input);
        allPainPoints.push(...batchResults.painPoints);
        modelCalls += batchResults.modelCalls;
        totalTokensUsed += batchResults.tokensUsed;
      }
      
      // Remove duplicates and merge similar pain points
      const uniquePainPoints = this.deduplicatePainPoints(allPainPoints);
      
      // Generate summary statistics
      const summaryStats = this.generateSummaryStatistics(uniquePainPoints, input.raw_feedback);
      
      // Calculate quality indicators
      const qualityIndicators = this.calculateQualityIndicators(uniquePainPoints, input.raw_feedback);
      
      const executionTime = Date.now() - startTime;
      
      console.log(`✅ ExtractionAgent: Completed analysis - ${uniquePainPoints.length} pain points found (${executionTime}ms)`);
      
      return {
        success: true,
        data: {
          pain_points: uniquePainPoints,
          summary_statistics: summaryStats,
          quality_indicators: qualityIndicators
        },
        execution_time_ms: executionTime,
        agent_name: 'ExtractionAgent',
        metadata: {
          model_calls: modelCalls,
          tokens_used: totalTokensUsed,
          confidence: qualityIndicators.extraction_confidence,
          debug_info: {
            execution_id: executionId,
            batch_count: Math.ceil(input.raw_feedback.length / batchSize),
            items_processed: input.raw_feedback.length,
            duplicates_removed: allPainPoints.length - uniquePainPoints.length
          }
        }
      };
      
    } catch (error: any) {
      const executionTime = Date.now() - startTime;
      console.error('❌ ExtractionAgent Error:', error);
      
      return {
        success: false,
        error: `Extraction failed: ${error.message}`,
        execution_time_ms: executionTime,
        agent_name: 'ExtractionAgent',
        metadata: {
          model_calls: 0,
          confidence: 0,
          debug_info: {
            execution_id: executionId,
            error_type: error.name || 'UnknownError',
            error_location: 'ExtractionAgent.execute'
          }
        }
      };
    }
  }

  /**
   * Process a batch of feedback items
   */
  private async processBatch(
    batch: RawFeedbackItem[], 
    input: ExtractionAgentInput
  ): Promise<{ painPoints: ExtractedPainPoint[]; modelCalls: number; tokensUsed: number }> {
    
    // Prepare batch data for analysis
    const batchData = batch.map(item => ({
      id: item.id,
      title: item.title,
      content: item.content.substring(0, 1000), // Limit content length
      source: item.source,
      metadata: {
        score: item.metadata.score,
        tags: item.metadata.tags?.slice(0, 5), // Limit tags
        created_date: item.metadata.created_date,
        repository: item.metadata.repository
      }
    }));
    
    // Build context-aware prompt
    const prompt = this.buildBatchPrompt(batchData, input.analysis_focus);
    
    // Execute AI analysis
    const result = await this.model.send({
    role: 'user',
    content: prompt
    });

    // Add null check for result.content
    const resultContent = result.content || '';

    // Parse AI response
    const painPoints = this.parseExtractionResult(resultContent, batch);
    
    return {
    painPoints,
    modelCalls: 1,
    tokensUsed: this.estimateTokens(prompt) + this.estimateTokens(resultContent)
    };
  }

  /**
   * Build the extraction prompt for a batch of feedback
   */
  private buildBatchPrompt(batchData: any[], analysisFocus?: any): string {
    const focusInstructions = analysisFocus ? `
Focus areas:
- Extract technical details: ${analysisFocus.extract_technical_details}
- Prioritize recent issues: ${analysisFocus.prioritize_recent}
- Group similar issues: ${analysisFocus.group_similar_issues}
` : '';

    return `${this.extractionPrompt}

${focusInstructions}

FEEDBACK DATA TO ANALYZE:
${JSON.stringify(batchData, null, 2)}

Extract pain points and return ONLY valid JSON array of pain point objects. Each object must have ALL required fields.`;
  }

  /**
   * Core extraction prompt template
   */
  private buildExtractionPrompt(): string {
    return `You are an expert analyst extracting developer pain points from Microsoft Teams Platform community feedback.

TASK: Transform raw feedback into structured pain point analysis with sentiment and classification.

OUTPUT FORMAT: Return ONLY a JSON array of pain point objects. Each object must include:

{
  "id": "unique_identifier",
  "primary_pain_point": "clear description of main issue",
  "category": "documentation" | "engineering" | "product",
  "priority": "escalation_risk" | "sentiment_spike" | "normal",
  "sentiment": "very_negative" | "negative" | "neutral" | "positive" | "very_positive",
  "confidence": 0.85,
  "affected_components": ["teams-sdk", "bot-framework", "adaptive-cards"],
  "technical_details": {
    "error_messages": ["specific error text"],
    "code_snippets": ["relevant code if present"],
    "environment_info": ["version numbers, OS, etc"]
  },
  "impact_assessment": {
    "developer_personas": ["beginner", "intermediate", "expert"],
    "blocking_severity": "critical" | "high" | "medium" | "low",
    "frequency": "widespread" | "common" | "occasional" | "rare"
  },
  "source_items": ["feedback_item_id_1", "feedback_item_id_2"]
}

ANALYSIS GUIDELINES:

Categories:
- "documentation": Missing docs, unclear guides, outdated examples, insufficient tutorials
- "engineering": Bugs, crashes, performance issues, API failures, SDK limitations  
- "product": Feature gaps, UX friction, platform limitations, missing capabilities

Priority Levels:
- "escalation_risk": Critical blocking issues, security concerns, widespread impact
- "sentiment_spike": High frustration, viral negative feedback, urgent community response needed
- "normal": Standard feedback, feature requests, general questions

Sentiment Analysis:
- Focus on developer frustration level and impact on productivity
- Consider urgency indicators ("urgent", "blocking", "can't proceed")
- Look for emotional language ("frustrated", "impossible", "terrible")

Technical Detail Extraction:
- Extract exact error messages and codes
- Identify version numbers, environment details
- Capture relevant code snippets or configuration issues
- Note specific APIs, methods, or components mentioned

Impact Assessment:
- Determine which developer skill levels are affected
- Assess how blocking the issue is for development workflows
- Evaluate how widespread the issue appears to be

IMPORTANT: Return ONLY the JSON array. No additional text, explanations, or markdown formatting.`;
  }

  /**
   * Parse AI extraction result and validate structure
   */
  private parseExtractionResult(content: string, originalItems: RawFeedbackItem[]): ExtractedPainPoint[] {
    try {
      // Clean the content - remove any markdown formatting
      const cleanContent = content.replace(/```json\n?|\n?```/g, '').trim();
      
      const parsed = JSON.parse(cleanContent);
      const painPoints = Array.isArray(parsed) ? parsed : [parsed];
      
      return painPoints.map((pp, index) => {
        // Validate and provide defaults for required fields
        const sourceItems = originalItems.map(item => item.id);
        
        return {
          id: pp.id || `extracted_${Date.now()}_${index}`,
          primary_pain_point: pp.primary_pain_point || 'Unable to extract pain point',
          category: this.validateCategory(pp.category),
          priority: this.validatePriority(pp.priority),
          sentiment: this.validateSentiment(pp.sentiment),
          confidence: Math.max(0, Math.min(1, pp.confidence || 0.5)),
          affected_components: Array.isArray(pp.affected_components) ? pp.affected_components : [],
          technical_details: {
            error_messages: pp.technical_details?.error_messages || [],
            code_snippets: pp.technical_details?.code_snippets || [],
            environment_info: pp.technical_details?.environment_info || []
          },
          impact_assessment: {
            developer_personas: Array.isArray(pp.impact_assessment?.developer_personas) 
              ? pp.impact_assessment.developer_personas 
              : ['intermediate'],
            blocking_severity: this.validateSeverity(pp.impact_assessment?.blocking_severity),
            frequency: this.validateFrequency(pp.impact_assessment?.frequency)
          },
          source_items: Array.isArray(pp.source_items) ? pp.source_items : sourceItems
        };
      });
      
    } catch (error) {
      console.warn('Failed to parse extraction result, creating fallback pain point');
      
      // Create fallback pain point from original data
      return [{
        id: `fallback_${Date.now()}`,
        primary_pain_point: originalItems[0]?.title || 'Failed to extract pain point',
        category: 'engineering',
        priority: 'normal',
        sentiment: 'neutral',
        confidence: 0.3,
        affected_components: originalItems[0]?.metadata.tags || [],
        technical_details: {
          error_messages: [],
          code_snippets: [],
          environment_info: []
        },
        impact_assessment: {
          developer_personas: ['intermediate'],
          blocking_severity: 'medium',
          frequency: 'occasional'
        },
        source_items: originalItems.map(item => item.id)
      }];
    }
  }

  /**
   * Remove duplicate pain points and merge similar ones
   */
  private deduplicatePainPoints(painPoints: ExtractedPainPoint[]): ExtractedPainPoint[] {
    const unique: ExtractedPainPoint[] = [];
    const seen = new Set<string>();
    
    for (const pp of painPoints) {
      // Create similarity key based on primary pain point and category
      const similarityKey = `${pp.category}_${pp.primary_pain_point.toLowerCase().substring(0, 50)}`;
      
      if (!seen.has(similarityKey)) {
        seen.add(similarityKey);
        unique.push(pp);
      } else {
        // Merge with existing similar pain point
        const existing = unique.find(u => 
          u.category === pp.category && 
          u.primary_pain_point.toLowerCase().substring(0, 50) === pp.primary_pain_point.toLowerCase().substring(0, 50)
        );
        
        if (existing) {
          // Merge source items and improve confidence
          existing.source_items = [...new Set([...existing.source_items, ...pp.source_items])];
          existing.confidence = Math.max(existing.confidence, pp.confidence);
          existing.affected_components = [...new Set([...existing.affected_components, ...pp.affected_components])];
        }
      }
    }
    
    return unique;
  }

  /**
   * Generate summary statistics
   */
  private generateSummaryStatistics(painPoints: ExtractedPainPoint[], rawFeedback: RawFeedbackItem[]) {
    const categoryBreakdown = painPoints.reduce((acc, pp) => {
      acc[pp.category] = (acc[pp.category] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);
    
    const priorityBreakdown = painPoints.reduce((acc, pp) => {
      acc[pp.priority] = (acc[pp.priority] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);
    
    const sentimentBreakdown = painPoints.reduce((acc, pp) => {
      acc[pp.sentiment] = (acc[pp.sentiment] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);
    
    return {
      total_items_analyzed: rawFeedback.length,
      pain_points_found: painPoints.length,
      category_breakdown: categoryBreakdown,
      priority_breakdown: priorityBreakdown,
      sentiment_breakdown: sentimentBreakdown
    };
  }

  /**
   * Calculate quality indicators
   */
  private calculateQualityIndicators(painPoints: ExtractedPainPoint[], rawFeedback: RawFeedbackItem[]) {
    const avgConfidence = painPoints.length > 0 
      ? painPoints.reduce((sum, pp) => sum + pp.confidence, 0) / painPoints.length 
      : 0;
    
    const dataCompleteness = painPoints.length > 0 
        ? painPoints.filter(pp => 
            (pp.technical_details?.error_messages?.length ?? 0) > 0 || 
            pp.affected_components.length > 0
        ).length / painPoints.length
        : 0;
        
    const categorizationCertainty = painPoints.length > 0
      ? painPoints.filter(pp => pp.confidence > 0.7).length / painPoints.length
      : 0;
    
    return {
      extraction_confidence: avgConfidence,
      data_completeness: dataCompleteness,
      categorization_certainty: categorizationCertainty
    };
  }

  // Validation helper methods
  private validateCategory(category: any): 'documentation' | 'engineering' | 'product' {
    const validCategories = ['documentation', 'engineering', 'product'];
    return validCategories.includes(category) ? category : 'engineering';
  }

  private validatePriority(priority: any): 'escalation_risk' | 'sentiment_spike' | 'normal' {
    const validPriorities = ['escalation_risk', 'sentiment_spike', 'normal'];
    return validPriorities.includes(priority) ? priority : 'normal';
  }

  private validateSentiment(sentiment: any): 'very_negative' | 'negative' | 'neutral' | 'positive' | 'very_positive' {
    const validSentiments = ['very_negative', 'negative', 'neutral', 'positive', 'very_positive'];
    return validSentiments.includes(sentiment) ? sentiment : 'neutral';
  }

  private validateSeverity(severity: any): 'critical' | 'high' | 'medium' | 'low' {
    const validSeverities = ['critical', 'high', 'medium', 'low'];
    return validSeverities.includes(severity) ? severity : 'medium';
  }

  private validateFrequency(frequency: any): 'widespread' | 'common' | 'occasional' | 'rare' {
    const validFrequencies = ['widespread', 'common', 'occasional', 'rare'];
    return validFrequencies.includes(frequency) ? frequency : 'occasional';
  }

  private validateInput(input: ExtractionAgentInput): void {
    if (!input.raw_feedback || input.raw_feedback.length === 0) {
      throw new Error('No feedback data provided for extraction');
    }
    
    if (input.raw_feedback.length > 50) {
      throw new Error('Too many feedback items - maximum 50 allowed per extraction');
    }
  }

  private estimateTokens(text: string): number {
    // Rough estimate: 1 token ≈ 4 characters
    return Math.ceil(text.length / 4);
  }
}