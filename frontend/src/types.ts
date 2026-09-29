export type PreparationStatus = 'pending_enqueue' | 'queued' | 'running' | 'ready' | 'failed' | 'not_started'

export interface Preparation {
  status: PreparationStatus
  stage?: string | null
  error?: string | null
  prepared_at?: string | null
}

export interface SystemInfo {
  model_mode: 'local_draft' | 'configured_api' | 'unknown'
  model_configured: boolean | null
  worker_status: 'reported' | 'unknown' | 'inline'
  text_model: string | null
  vision_model: string | null
  reported_at: string | null
}

export interface Account {
  id: string
  email: string
}

export interface Project {
  id: string
  name: string
  created_at: string
}

export interface Source {
  id: string
  project_id: string
  kind: 'template' | 'content'
  filename: string
  sha256: string
  size_bytes: number
  created_at: string
  preparation?: Preparation | null
}

export type RunStatus =
  | 'pending_enqueue'
  | 'queued'
  | 'running'
  | 'awaiting_selection'
  | 'completed'
  | 'completed_with_warnings'
  | 'failed'
  | 'interrupted'
  | 'cancelled'

export interface Run {
  id: string
  project_id: string
  kind: string
  status: RunStatus
  stage?: string | null
  progress?: number | null
  created_at: string
  error?: string | null
  warnings?: string[]
  durations?: Record<string, number | boolean | null>
  artifacts?: Artifact[]
  versions?: string[]
}

export interface Artifact {
  id: string
  kind: string
  url: string
}

export interface Version {
  id: string
  project_id: string
  run_id: string
  parent_version_id?: string | null
  variant_id?: string | null
  ordinal: number
  quality_status: string
  plan?: Record<string, unknown> | null
  created_at: string
  preview_urls: string[]
  artifacts: Artifact[]
}

export interface AuditIssue {
  id: string
  version_id: string
  slide_id?: string | null
  object_id?: string | null
  rule_id: string
  severity: string
  evidence: string | Record<string, unknown>
  selected: boolean
  repairability?: 'automatic' | 'manual'
  resolution?: string | null
}
