// src/types/index.ts - Foundation interfaces for agent chain
/* eslint-disable @typescript-eslint/no-explicit-any */

// ============================================================================
// CORE DATA STRUCTURES
// ============================================================================

export interface RawFeedbackItem {
  id: string;
  title: string;
  content: string;
  url: string;
  source: 'stackoverflow' | 'github';
  metadata: {
    score?: number;
    tags?: string[];
    author?: string;
    created_date: string;
    repository?: string;
    labels?: string[];
  };
}

export interface ExtractedPainPoint {
  id: string;
  primary_pain_point: string;
  category: 'documentation' | 'engineering' | 'product';
  priority: 'escalation_risk' | 'sentiment_spike' | 'normal';
  sentiment: 'very_negative' | 'negative' | 'neutral' | 'positive' | 'very_positive';
  confidence: number; // 0.0 to 1.0
  affected_components: string[];
  technical_details: {
    error_messages?: string[];
    code_snippets?: string[];
    environment_info?: string[];
  };
  impact_assessment: {
    developer_personas: ('beginner' | 'intermediate' | 'expert')[];
    blocking_severity: 'critical' | 'high' | 'medium' | 'low';
    frequency: 'widespread' | 'common' | 'occasional' | 'rare';
  };
  source_items: string[]; // IDs of raw feedback items that contributed
}

export interface CommunityResponse {
  response_text: string;
  tone: 'professional' | 'friendly' | 'technical';
  includes_workaround: boolean;
  includes_timeline: boolean;
  tracking_references: string[];
  confidence_level: number;
  estimated_helpfulness: number;
}

// ============================================================================
// AGENT INTERFACES
// ============================================================================

export interface AgentExecutionContext {
  request_id: string;
  user_query: string;
  timestamp: string;
  conversation_history?: ConversationTurn[];
  debug_mode?: boolean;
}

export interface AgentResult<T> {
  success: boolean;
  data?: T;
  error?: string;
  execution_time_ms: number;
  agent_name: string;
  metadata: {
    model_calls: number;
    tokens_used?: number;
    confidence: number;
    debug_info?: any;
  };
}

// ============================================================================
// EXTRACTION AGENT TYPES
// ============================================================================

export interface ExtractionAgentInput {
  raw_feedback: RawFeedbackItem[];
  analysis_focus?: {
    extract_technical_details: boolean;
    prioritize_recent: boolean;
    group_similar_issues: boolean;
  };
  context: AgentExecutionContext;
}

export interface ExtractionAgentOutput {
  pain_points: ExtractedPainPoint[];
  summary_statistics: {
    total_items_analyzed: number;
    pain_points_found: number;
    category_breakdown: Record<string, number>;
    priority_breakdown: Record<string, number>;
    sentiment_breakdown: Record<string, number>;
  };
  quality_indicators: {
    extraction_confidence: number;
    data_completeness: number;
    categorization_certainty: number;
  };
}

// ============================================================================
// RESPONSE AGENT TYPES
// ============================================================================

export interface ResponseAgentInput {
  pain_points: ExtractedPainPoint[];
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
  context: AgentExecutionContext;
}

export interface ResponseAgentOutput {
  community_response: CommunityResponse;
  response_metadata: {
    knowledge_base_matches: string[];
    template_used: string;
    personalization_applied: boolean;
    estimated_user_satisfaction: number;
  };
  suggested_actions: Array<{
    action_type: 'post_response' | 'escalate_internal' | 'update_documentation' | 'create_tracking_issue';
    action_label: string;
    action_data: any;
    priority: 'high' | 'medium' | 'low';
  }>;
}

// ============================================================================
// AGENT CHAIN ORCHESTRATION
// ============================================================================

export interface AgentChainConfig {
  extraction_agent: {
    model_config: {
      temperature: number;
      max_tokens: number;
      timeout_ms: number;
    };
    prompt_template: string;
    validation_rules: string[];
  };
  response_agent: {
    model_config: {
      temperature: number;
      max_tokens: number;
      timeout_ms: number;
    };
    knowledge_base_enabled: boolean;
    response_templates_enabled: boolean;
  };
  error_handling: {
    max_retries: number;
    fallback_responses: Record<string, string>;
    escalation_threshold: number;
  };
}

export interface AgentChainResult {
  extraction_result: AgentResult<ExtractionAgentOutput>;
  response_result: AgentResult<ResponseAgentOutput>;
  overall_success: boolean;
  total_execution_time_ms: number;
  chain_metadata: {
    steps_completed: number;
    steps_failed: number;
    quality_score: number;
    user_experience_score: number;
  };
}

// ============================================================================
// PROGRESS TRACKING
// ============================================================================

export interface ProgressUpdate {
  stage: 'ingestion' | 'extraction' | 'response_generation' | 'complete';
  message: string;
  progress_percentage: number;
  estimated_time_remaining_ms?: number;
  current_operation?: string;
}

export type ProgressCallback = (update: ProgressUpdate) => Promise<void>;

// ============================================================================
// CONVERSATION MEMORY
// ============================================================================

export interface ConversationTurn {
  turn_id: string;
  user_input: string;
  bot_response: string;
  timestamp: string;
  analysis_performed?: {
    platforms_analyzed: string[];
    pain_points_found: number;
    response_generated: boolean;
  };
  user_actions?: Array<{
    action_type: string;
    action_data: any;
    timestamp: string;
  }>;
}

export interface ConversationMemory {
  conversation_id: string;
  user_id: string;
  turns: ConversationTurn[];
  persistent_context: {
    preferred_analysis_style?: 'detailed' | 'summary' | 'technical';
    frequently_analyzed_topics: string[];
    pain_point_history: ExtractedPainPoint[];
  };
  created_at: string;
  last_updated: string;
}

// ============================================================================
// ADAPTIVE CARD TYPES
// ============================================================================

export interface CardAction {
  type: 'post_response' | 'escalate' | 'track_issue' | 'update_docs' | 'retry_analysis';
  label: string;
  enabled: boolean;
  action_data: any;
  style?: 'default' | 'positive' | 'destructive';
}

export interface InsightsCardData {
  analysis_summary: {
    platform: string;
    query: string;
    total_items: number;
    pain_points_count: number;
  };
  priority_insights: Array<{
    title: string;
    category: string;
    priority: string;
    sentiment: string;
    description: string;
  }>;
  generated_response?: {
    text: string;
    confidence: number;
    includes_workaround: boolean;
  };
  actions: CardAction[];
  visual_styling: {
    color_theme: 'success' | 'warning' | 'attention' | 'accent';
    priority_indicators: boolean;
    progress_bars: boolean;
  };
}

// ============================================================================
// EVALUATION & OBSERVABILITY
// ============================================================================

export interface PerformanceMetrics {
  operation_id: string;
  operation_type: 'extraction' | 'response_generation' | 'full_chain';
  start_time: string;
  end_time: string;
  duration_ms: number;
  success: boolean;
  error_type?: string;
  resource_usage: {
    model_calls: number;
    tokens_consumed: number;
    memory_peak_mb: number;
  };
  quality_metrics: {
    output_relevance: number;
    output_completeness: number;
    user_satisfaction_predicted: number;
  };
}

export interface EvaluationCriteria {
  accuracy_weight: number;
  relevance_weight: number;
  helpfulness_weight: number;
  response_time_weight: number;
  user_satisfaction_weight: number;
}

export interface EvaluationResult {
  overall_score: number;
  component_scores: {
    accuracy: number;
    relevance: number;
    helpfulness: number;
    response_time: number;
    user_satisfaction: number;
  };
  improvement_suggestions: string[];
  benchmark_comparison: {
    vs_previous_version: number;
    vs_baseline: number;
  };
}

// ============================================================================
// ERROR HANDLING
// ============================================================================

export interface ProcessingError {
  error_id: string;
  error_type: 'tool_failure' | 'model_timeout' | 'validation_failed' | 'rate_limit' | 'unknown';
  error_message: string;
  component: 'stackoverflow_tool' | 'github_tool' | 'extraction_agent' | 'response_agent' | 'orchestrator';
  timestamp: string;
  context: any;
  recovery_attempted: boolean;
  user_impact: 'blocking' | 'degraded' | 'minimal';
}

export interface ErrorRecoveryStrategy {
  error_type: string;
  recovery_actions: Array<{
    action: string;
    timeout_ms: number;
    fallback_on_failure?: string;
  }>;
  user_communication: {
    show_progress: boolean;
    message_template: string;
    include_eta: boolean;
  };
}

// ============================================================================
// CONFIGURATION TYPES
// ============================================================================

export interface EnvironmentConfig {
  azure_openai: {
    endpoint: string;
    api_key: string;
    model: string;
    api_version: string;
  };
  github: {
    token: string;
    demo_issue_url: string;
  };
  stackoverflow: {
    api_key?: string;
    rate_limit_ms: number;
  };
  features: {
    conversation_memory_enabled: boolean;
    progress_updates_enabled: boolean;
    knowledge_base_enabled: boolean;
    evaluation_enabled: boolean;
  };
}

// ============================================================================
// UTILITY TYPES
// ============================================================================

export type Platform = 'stackoverflow' | 'github';
export type AnalysisType = 'stackoverflow_analysis' | 'github_analysis' | 'combined_analysis';
export type PainPointCategory = 'documentation' | 'engineering' | 'product';
export type PriorityLevel = 'escalation_risk' | 'sentiment_spike' | 'normal';
export type SentimentLevel = 'very_negative' | 'negative' | 'neutral' | 'positive' | 'very_positive';

// ============================================================================
// EXPORTS
// ============================================================================

export * from './index';