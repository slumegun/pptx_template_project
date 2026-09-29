import { useCallback, useEffect, useMemo, useState } from 'react'
import { api, downloadFile } from './api'
import type { Account, AuditIssue, Project, Run, Source, Version } from './types'
import HomePage from './HomePage'
import type { AuthMode } from './HomePage'
import AuthPage from './AuthPage'
import CreatePage from './CreatePage'
import ProgressPage from './ProgressPage'
import ComparePage from './ComparePage'
import AuditPage from './AuditPage'
import { Icon } from './icons'

type Screen = 'home' | 'auth' | 'create' | 'progress' | 'compare' | 'audit'
const terminal = new Set(['completed', 'completed_with_warnings', 'awaiting_selection', 'failed', 'interrupted', 'cancelled'])

export default function App() {
  const [screen, setScreen] = useState<Screen>('home')
  const [authMode, setAuthMode] = useState<AuthMode>('login')
  const [, setAccount] = useState<Account | null>(null)
  const [authBusy, setAuthBusy] = useState(false)
  const [project, setProject] = useState<Project | null>(null)
  const [templateId, setTemplateId] = useState<string | null>(null)
  const [template, setTemplate] = useState<Source | null>(null)
  const [templates, setTemplates] = useState<Source[]>([])
  const [brief, setBrief] = useState('')
  const [slideCount, setSlideCount] = useState(12)
  const [runId, setRunId] = useState<string | null>(null)
  const [run, setRun] = useState<Run | null>(null)
  const [generationId, setGenerationId] = useState<string | null>(null)
  const [versions, setVersions] = useState<Version[]>([])
  const [selectedVersionId, setSelectedVersionId] = useState<string | null>(null)
  const [issues, setIssues] = useState<AuditIssue[]>([])
  const [selectedIssueIds, setSelectedIssueIds] = useState<string[]>([])
  const [busy, setBusy] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const [theme, setTheme] = useState<'dark' | 'light'>(() => localStorage.getItem('lukas-theme') === 'light' ? 'light' : 'dark')

  useEffect(() => { localStorage.setItem('lukas-theme', theme) }, [theme])

  const refreshVersions = useCallback(async (projectId: string) => {
    const next = await api.getVersions(projectId)
    setVersions(Array.isArray(next) ? next : [])
    return next
  }, [])

  useEffect(() => {
    let active = true
    void api.me().then(async user => {
      if (!active) return
      await loadWorkspace(user, () => active)
    }).catch(() => { /* A guest stays on the landing page. */ })
    return () => { active = false }
  }, [])

  async function loadWorkspace(user: Account, isActive = () => true) {
    try {
      const nextProject = await api.workspace()
      const [sources, nextVersions, runs] = await Promise.all([
        api.getProjectSources(nextProject.id),
        api.getVersions(nextProject.id),
        api.getRuns(nextProject.id),
      ])
      if (!isActive()) return
      const nextTemplates = sources.filter(source => source.kind === 'template')
      setAccount(user)
      setProject(nextProject)
      setTemplates(nextTemplates)
      setTemplate(nextTemplates[0] || null)
      setTemplateId(nextTemplates[0]?.id || null)
      setVersions(Array.isArray(nextVersions) ? nextVersions : [])
      const restored = runs.find(item => !terminal.has(item.status)) || runs[0] || null
      setRun(restored); setRunId(restored?.id || null)
      setGenerationId(null)
      setScreen(restored && !terminal.has(restored.status) ? 'progress' : 'create')
    } catch (e) {
      if (isActive()) { setError(message(e)); setScreen('auth') }
    }
  }

  async function authenticate(mode: AuthMode, email: string, password: string) {
    setAuthBusy(true); setError(null)
    try {
      const user = mode === 'login' ? await api.login(email, password) : await api.register(email, password)
      await loadWorkspace(user)
    } catch (e) { setError(message(e)) }
    finally { setAuthBusy(false) }
  }

  async function signOut() {
    try { await api.logout() } catch { /* Clear local state even if the server is unavailable. */ }
    setAccount(null); setProject(null); setTemplates([]); setTemplate(null); setTemplateId(null)
    setVersions([]); setRun(null); setRunId(null); setSelectedVersionId(null)
    setGenerationId(null); setBrief(''); setScreen('home'); setSidebarOpen(false); setError(null)
  }

  function openAuth(mode: AuthMode) { setAuthMode(mode); setError(null); setScreen('auth') }

  useEffect(() => {
    if (!project) return
    void api.getProjectSources(project.id)
      .then(sources => setTemplates(sources.filter(source => source.kind === 'template')))
      .catch(e => setError(message(e)))
  }, [project?.id])
  useEffect(() => {
    if (!project) return
    void refreshVersions(project.id).catch(e => setError(message(e)))
  }, [project?.id, refreshVersions])

  const preparingTemplates = templates.some(source => !['ready', 'failed'].includes(source.preparation?.status || 'queued'))
  useEffect(() => {
    if (!project) return
    let disposed = false
    let inFlight = false
    async function check() {
      if (inFlight) return
      inFlight = true
      try {
        const sources = await api.getProjectSources(project!.id)
        if (disposed) return
        const nextTemplates = sources.filter(source => source.kind === 'template')
        setTemplates(nextTemplates)
        setTemplate(previous => nextTemplates.find(source => source.id === previous?.id) || previous)
      } catch (e) { if (!disposed) setError(message(e)) }
      finally { inFlight = false }
    }
    void check()
    if (!preparingTemplates) return () => { disposed = true }
    const timer = window.setInterval(check, 2000)
    return () => { disposed = true; window.clearInterval(timer) }
  }, [project?.id, preparingTemplates, templateId])

  useEffect(() => {
    if (!runId) return
    let disposed = false
    let inFlight = false
    async function check() {
      if (inFlight) return
      inFlight = true
      try {
        const current = await api.getRun(runId!)
        if (disposed) return
        setRun(current)
        if (terminal.has(current.status) && project) void refreshVersions(project.id).catch(e => setError(message(e)))
      } catch (e) { if (!disposed) setError(message(e)) }
      finally { inFlight = false }
    }
    void check()
    if (run && terminal.has(run.status)) return () => { disposed = true }
    const timer = window.setInterval(check, 1600)
    return () => { disposed = true; window.clearInterval(timer) }
  }, [runId, run?.status, project?.id, refreshVersions])

  useEffect(() => {
    if (!selectedVersionId) { setIssues([]); return }
    let disposed = false
    Promise.all([api.getVersion(selectedVersionId), api.getIssues(selectedVersionId)]).then(([version, nextIssues]) => {
      if (disposed) return
      setVersions(previous => previous.some(item => item.id === version.id) ? previous.map(item => item.id === version.id ? version : item) : [version, ...previous])
      setIssues(Array.isArray(nextIssues) ? nextIssues : [])
    }).catch(e => { if (!disposed) setError(message(e)) })
    return () => { disposed = true }
  }, [selectedVersionId])

  const selectedVersion = versions.find(version => version.id === selectedVersionId) || null
  const compareVersions = useMemo(() => {
    const originals = versions.filter(version => !version.parent_version_id)
    if (!originals.length) return []
    const latest = [...originals].sort((a, b) => new Date(b.created_at).getTime() - new Date(a.created_at).getTime())[0]
    const roots = originals.filter(version => version.run_id === (generationId || latest.run_id)).sort((a, b) => a.ordinal - b.ordinal)
    const byId = new Map(versions.map(version => [version.id, version]))
    function rootId(version: Version): string {
      let current = version
      const seen = new Set<string>()
      while (current.parent_version_id && byId.has(current.parent_version_id) && !seen.has(current.id)) {
        seen.add(current.id)
        current = byId.get(current.parent_version_id)!
      }
      return current.id
    }
    return roots.map(root => versions.filter(version => rootId(version) === root.id)
      .sort((a, b) => new Date(b.created_at).getTime() - new Date(a.created_at).getTime())[0] || root)
  }, [versions, generationId])

  async function uploadTemplate(file: File) {
    if (!file.name.toLowerCase().endsWith('.pptx')) { setError('Для шаблона нужен файл PPTX.'); return }
    if (!project) return
    setBusy('template'); setError(null)
    try {
      const next = await api.uploadSource(project.id, file, 'template')
      setTemplateId(next.id); setTemplate(next); setTemplates(previous => [next, ...previous])
    } catch (e) { setError(message(e)) } finally { setBusy(null) }
  }

  function selectTemplate(source: Source) { setTemplateId(source.id); setTemplate(source) }

  async function importTemplate(file: File) {
    if (!file.name.toLowerCase().endsWith('.zip')) { setError('Выберите пакет шаблона .zip, скачанный кнопкой «Скачать».'); return }
    if (!project) return
    setBusy('import'); setError(null)
    try {
      const next = await api.importTemplate(project.id, file)
      setTemplateId(next.id); setTemplate(next)
      setTemplates(previous => [next, ...previous.filter(item => item.id !== next.id)])
    } catch (e) { setError(message(e)) } finally { setBusy(null) }
  }

  async function exportTemplate(source: Source) {
    if (!project) return
    setBusy('export'); setError(null)
    try {
      await downloadFile(api.templateExportUrl(project.id, source.id), source.filename.replace(/\.pptx$/i, '') + '.template.zip')
    } catch (e) { setError(message(e)) } finally { setBusy(null) }
  }


  async function deleteTemplate(source: Source) {
    if (!project) return
    setBusy('delete'); setError(null)
    try {
      await api.deleteTemplate(project.id, source.id)
      const remaining = templates.filter(item => item.id !== source.id)
      setTemplates(remaining)
      if (templateId === source.id) {
        setTemplate(remaining[0] || null)
        setTemplateId(remaining[0]?.id || null)
      }
    } catch (e) { setError(message(e)) }
    finally { setBusy(null) }
  }

  async function retryPreparation() {
    if (!project || !templateId) return
    setBusy('prepare'); setError(null)
    try {
      const source = await api.prepareSource(project.id, templateId)
      setTemplate(source)
      setTemplates(previous => previous.map(item => item.id === source.id ? source : item))
    } catch (e) { setError(message(e)) } finally { setBusy(null) }
  }

  async function generate() {
    if (!project || !template || template.preparation?.status !== 'ready') { setError('Дождитесь подготовки шаблона.'); return }
    if (!brief.trim()) { setError('Напишите задание для презентации.'); return }
    if (!Number.isInteger(slideCount) || slideCount < 1 || slideCount > 50) { setError('Укажите от 1 до 50 слайдов.'); return }
    setBusy('generate'); setError(null)
    try {
      const next = await api.createRun(project.id, { brief: brief.trim(), slide_count: slideCount, template_source_id: template.id, content_source_ids: [] })
      setRun(next); setRunId(next.id); setGenerationId(null); setScreen('progress')
    } catch (e) { setError(message(e)) } finally { setBusy(null) }
  }

  async function cancelRun() {
    if (!run) return
    setBusy('cancel'); setError(null)
    try { setRun(await api.cancelRun(run.id)) } catch (e) { setError(message(e)) } finally { setBusy(null) }
  }

  async function retryRun() {
    if (!run) return
    setBusy('retry'); setError(null)
    try { const next = await api.retryRun(run.id); setRun(next); setRunId(next.id) } catch (e) { setError(message(e)) } finally { setBusy(null) }
  }

  async function repair() {
    const automaticIssueIds = selectedIssueIds.filter(id => issues.find(issue => issue.id === id)?.repairability !== 'manual')
    if (!selectedVersion || !automaticIssueIds.length) return
    setBusy('repair'); setError(null)
    try {
      const next = await api.createRepair(selectedVersion.id, automaticIssueIds)
      setRun(next); setRunId(next.id); setSelectedIssueIds([]); setScreen('progress')
    } catch (e) { setError(message(e)) } finally { setBusy(null) }
  }

  async function editVersion(versionId: string, slideIndex: number, prompt: string): Promise<boolean> {
    setBusy('edit'); setError(null)
    try {
      const next = await api.createSlideEdit(versionId, slideIndex, prompt)
      setRun(next); setRunId(next.id); setScreen('progress')
      return true
    } catch (e) { setError(message(e)); return false }
    finally { setBusy(null) }
  }

  function selectVersion(id: string) { setSelectedVersionId(id); setSelectedIssueIds([]) }
  function openAudit(id: string) { selectVersion(id); navigate('audit') }
  function toggleIssue(id: string) { setSelectedIssueIds(previous => previous.includes(id) ? previous.filter(value => value !== id) : [...previous, id]) }
  function navigate(next: Screen) { setScreen(next); setSidebarOpen(false); setError(null) }
  async function openResults() {
    if (!project) return
    try {
      const currentVersions = await refreshVersions(project.id)
      let root = currentVersions.find(version => version.id === run?.versions?.[0])
      const seen = new Set<string>()
      while (root?.parent_version_id && !seen.has(root.id)) {
        seen.add(root.id)
        const parent = currentVersions.find(version => version.id === root!.parent_version_id)
        if (!parent) break
        root = parent
      }
      setGenerationId(root?.run_id || (run?.kind === 'generation' ? run.id : null))
      navigate('compare')
    } catch (e) { setError(message(e)) }
  }
  function newPresentation() { setBrief(''); setSelectedIssueIds([]); navigate('create') }
  const generationHistory = Array.from(new Map(versions.filter(version => !version.parent_version_id)
    .map(version => [version.run_id, version])).values())

  const navItems: { id: 'create' | 'compare'; label: string }[] = [
    { id: 'create', label: 'Шаблоны' },
    { id: 'compare', label: 'Презентации' },
  ]
  function openTab(id: 'create' | 'compare') {
    navigate(id === 'compare' && run && !terminal.has(run.status) ? 'progress' : id)
  }

  if (screen === 'home') return <HomePage onAuth={openAuth} />
  if (screen === 'auth') return <AuthPage mode={authMode} onMode={mode => { setAuthMode(mode); setError(null) }} onBack={() => navigate('home')} onSubmit={authenticate} busy={authBusy} error={error} />

  return <div className="app-shell" data-theme={theme}>
    <aside className={sidebarOpen ? 'sidebar sidebar--open' : 'sidebar'}>
      <div className="sidebar__top"><button type="button" className="brand" onClick={() => openTab('create')} aria-label="Лукас — шаблоны"><span className="brand__text"><strong>Лукас<span className="brand__dot">.</span></strong></span></button><button type="button" className="sidebar__close" onClick={() => setSidebarOpen(false)} aria-label="Закрыть меню"><Icon name="close" /></button></div>
      <div className="sidebar__group"><nav aria-label="Основная навигация">{navItems.map(item => <button type="button" key={item.id} className={'nav-item ' + ((screen === item.id || (item.id === 'compare' && screen === 'progress')) ? 'nav-item--active' : '')} onClick={() => openTab(item.id)}><span>{item.label}</span></button>)}</nav></div>
      <div className="sidebar__bottom"><div className="theme-switch" role="group" aria-label="Тема оформления"><button type="button" className={theme === 'dark' ? 'theme-switch__button theme-switch__button--active' : 'theme-switch__button'} onClick={() => setTheme('dark')} aria-pressed={theme === 'dark'}>Тёмная</button><button type="button" className={theme === 'light' ? 'theme-switch__button theme-switch__button--active' : 'theme-switch__button'} onClick={() => setTheme('light')} aria-pressed={theme === 'light'}>Светлая</button></div><button type="button" className="sidebar__logout" onClick={() => void signOut()}>Выйти</button></div>
    </aside>
    {sidebarOpen && <button type="button" className="sidebar-backdrop" onClick={() => setSidebarOpen(false)} aria-label="Закрыть меню" />}
    <div className="workspace">
      <div className="workspace__vk-symbol" aria-hidden="true" />
      <button type="button" className="workspace__mobile-menu" onClick={() => setSidebarOpen(true)} aria-label="Открыть меню"><Icon name="menu" /></button>
      <main>{error && <div className="global-error" role="alert"><Icon name="warning" size={19} /><span>{error}</span><button type="button" onClick={() => setError(null)} aria-label="Закрыть сообщение"><Icon name="close" size={16} /></button></div>}
        {screen === 'create' && <CreatePage project={project} template={template} templates={templates} onSelectTemplate={selectTemplate} onDeleteTemplate={deleteTemplate} onExportTemplate={exportTemplate} onImportTemplate={importTemplate} brief={brief} slideCount={slideCount} busy={busy} onTemplateFile={uploadTemplate} onBriefChange={setBrief} onSlideCountChange={setSlideCount} onRetryPreparation={retryPreparation} onGenerate={generate} />}
        {screen === 'progress' && <ProgressPage run={run} busy={busy} onCancel={cancelRun} onRetry={retryRun} onResults={openResults} onClose={() => navigate('create')} />}
        {screen === 'compare' && <>
          {generationHistory.length > 1 && <div className="generation-history">
            <label className="field-label" htmlFor="generation-history">История презентаций</label>
            <select id="generation-history" value={generationId || generationHistory[0]?.run_id || ''}
              onChange={event => { setGenerationId(event.target.value); setSelectedVersionId(null) }}>
              {generationHistory.map(version => <option key={version.run_id} value={version.run_id}>
                {new Date(version.created_at).toLocaleString('ru-RU')}
              </option>)}
            </select>
          </div>}
          <ComparePage key={generationId || generationHistory[0]?.run_id || 'empty'} versions={compareVersions}
            selectedId={selectedVersionId} onSelect={selectVersion} onAudit={openAudit} onEdit={editVersion}
            onNew={newPresentation} onError={setError} />
        </>}
        {screen === 'audit' && <AuditPage version={selectedVersion} issues={issues} selectedIssueIds={selectedIssueIds} busy={busy} onToggleIssue={toggleIssue} onRepair={repair} onBack={() => navigate('compare')} onError={setError} />}
      </main>
    </div>
  </div>
}

function message(error: unknown) { return error instanceof Error ? error.message : 'Произошла ошибка. Повторите попытку.' }
