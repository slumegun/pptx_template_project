import type { Account, Artifact, AuditIssue, Project, Run, Source, SystemInfo, Version } from './types'

const API_BASE = (import.meta.env.VITE_API_BASE_URL || '/api').replace(/\/$/, '')

export class ApiError extends Error {
  status: number
  constructor(message: string, status: number) {
    super(message)
    this.name = 'ApiError'
    this.status = status
  }
}

function extractMessage(body: unknown, fallback: string): string {
  if (typeof body === 'string' && body.trim()) return body
  if (body && typeof body === 'object' && 'detail' in body) {
    const detail = (body as { detail: unknown }).detail
    if (typeof detail === 'string') return detail
    if (Array.isArray(detail)) return detail.map(item => item?.msg || String(item)).join('; ')
  }
  return fallback
}

async function request<T>(path: string, options: RequestInit = {}): Promise<T> {
  const controller = new AbortController()
  const timeout = window.setTimeout(() => controller.abort(), options.body instanceof FormData ? 120000 : 30000)
  const abort = () => controller.abort()
  options.signal?.addEventListener('abort', abort, { once: true })
  if (options.signal?.aborted) controller.abort()
  try {
    const response = await fetch(`${API_BASE}${path}`, { ...options, signal: controller.signal })
    if (!response.ok) {
      let body: unknown = null
      try { body = await response.json() } catch { /* server may send plain text */ }
      throw new ApiError(extractMessage(body, `Сервер вернул ошибку ${response.status}.`), response.status)
    }
    if (response.status === 204) return undefined as T
    return await response.json() as T
  } catch (error) {
    if (error instanceof ApiError) throw error
    throw new ApiError(controller.signal.aborted
      ? 'Сервер не ответил вовремя. Проверьте соединение. Статус генерации обновится после восстановления связи.'
      : 'Не удалось связаться с сервером. Проверьте подключение и повторите попытку.', 0)
  } finally {
    window.clearTimeout(timeout)
    options.signal?.removeEventListener('abort', abort)
  }
}

function json(body: unknown): RequestInit {
  return { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }
}

export const api = {
  register(email: string, password: string) {
    return request<Account>('/auth/register', json({ email, password }))
  },
  login(email: string, password: string) {
    return request<Account>('/auth/login', json({ email, password }))
  },
  me() {
    return request<Account>('/auth/me')
  },
  logout() {
    return request<void>('/auth/logout', { method: 'POST' })
  },
  workspace() {
    return request<Project>('/workspace')
  },
  getSystem() {
    return request<SystemInfo>('/system')
  },
  createProject(name: string) {
    return request<Project>('/projects', json({ name }))
  },
  getProjects() {
    return request<Project[]>('/projects')
  },
  getProjectSources(projectId: string) {
    return request<Source[]>(`/projects/${encodeURIComponent(projectId)}/sources`)
  },
  uploadSource(projectId: string, file: File, kind: Source['kind']) {
    const body = new FormData()
    body.append('file', file)
    body.append('kind', kind)
    return request<Source>(`/projects/${encodeURIComponent(projectId)}/sources`, { method: 'POST', body })
  },
  importTemplate(projectId: string, file: File) {
    const body = new FormData()
    body.append('file', file)
    return request<Source>(`/projects/${encodeURIComponent(projectId)}/sources/import`, { method: 'POST', body })
  },
  templateExportUrl(projectId: string, sourceId: string) {
    return `${API_BASE}/projects/${encodeURIComponent(projectId)}/sources/${encodeURIComponent(sourceId)}/export`
  },
  deleteTemplate(projectId: string, sourceId: string) {
    return request<void>('/projects/' + encodeURIComponent(projectId) + '/sources/' + encodeURIComponent(sourceId), { method: 'DELETE' })
  },
  getSource(sourceId: string) {
    return request<Source>(`/sources/${encodeURIComponent(sourceId)}`)
  },
  prepareSource(projectId: string, sourceId: string) {
    return request<Source>(
      `/projects/${encodeURIComponent(projectId)}/sources/${encodeURIComponent(sourceId)}/prepare`,
      { method: 'POST' },
    )
  },
  createRun(projectId: string, payload: {
    brief: string
    slide_count: number
    template_source_id: string
    content_source_ids: string[]
  }) {
    return request<Run>(`/projects/${encodeURIComponent(projectId)}/runs`, json(payload))
  },
  getRuns(projectId: string) {
    return request<Run[]>(`/projects/${encodeURIComponent(projectId)}/runs`)
  },
  getRun(runId: string) {
    return request<Run>(`/runs/${encodeURIComponent(runId)}`)
  },
  cancelRun(runId: string) {
    return request<Run>(`/runs/${encodeURIComponent(runId)}/cancel`, { method: 'POST' })
  },
  retryRun(runId: string) {
    return request<Run>(`/runs/${encodeURIComponent(runId)}/retry`, { method: 'POST' })
  },
  getVersions(projectId: string) {
    return request<Version[]>(`/projects/${encodeURIComponent(projectId)}/versions`)
  },
  getVersion(versionId: string) {
    return request<Version>(`/versions/${encodeURIComponent(versionId)}`)
  },
  getIssues(versionId: string) {
    return request<AuditIssue[]>(`/versions/${encodeURIComponent(versionId)}/issues`)
  },
  createSlideEdit(versionId: string, slideIndex: number, prompt: string) {
    return request<Run>('/versions/' + encodeURIComponent(versionId) + '/edits', json({ slide_index: slideIndex, prompt }))
  },
  createRepair(versionId: string, issueIds: string[]) {
    return request<Run>(`/versions/${encodeURIComponent(versionId)}/repairs`, json({ issue_ids: issueIds }))
  },
  artifactUrl(artifact: Artifact) {
    return resolveUrl(artifact.url || `/api/artifacts/${artifact.id}`)
  },
}

export function resolveUrl(value: string): string {
  if (/^https?:\/\//i.test(value) || value.startsWith('blob:') || value.startsWith('data:')) return value
  if (value.startsWith('/')) return value
  return `${API_BASE}/${value.replace(/^\//, '')}`
}

export async function downloadArtifact(artifact: Artifact, filename: string): Promise<void> {
  await downloadFile(api.artifactUrl(artifact), filename, false)
}

/** Saves a file; the server's name wins unless the caller names the file itself. */
export async function downloadFile(url: string, fallbackName: string, serverName = true): Promise<void> {
  let response: Response
  try {
    response = await fetch(url)
  } catch {
    throw new ApiError('Не удалось скачать файл. Проверьте подключение.', 0)
  }
  if (!response.ok) {
    let body: unknown = null
    try { body = await response.json() } catch { /* a file error may have no JSON body */ }
    throw new ApiError(extractMessage(body, 'Не удалось скачать файл. Повторите попытку.'), response.status)
  }
  const disposition = response.headers.get('Content-Disposition') || ''
  const encoded = /filename\*=UTF-8''([^;]+)/i.exec(disposition)?.[1]
  const named = encoded ? decodeURIComponent(encoded) : /filename="([^"]+)"/i.exec(disposition)?.[1]
  const filename = serverName && named ? named : fallbackName
  const blob = await response.blob()
  const objectUrl = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = objectUrl
  link.download = filename
  document.body.appendChild(link)
  link.click()
  link.remove()
  window.setTimeout(() => URL.revokeObjectURL(objectUrl), 30_000)
}
