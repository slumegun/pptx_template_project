import { useEffect, useRef, useState } from 'react'
import type { ChangeEvent, DragEvent, ReactNode } from 'react'
import { downloadArtifact, resolveUrl } from './api'
import type { Artifact, Version } from './types'
import { Icon } from './icons'

export function Button({ children, variant = 'primary', icon, className = '', ...props }: React.ButtonHTMLAttributes<HTMLButtonElement> & { variant?: 'primary' | 'secondary' | 'ghost' | 'danger'; icon?: ReactNode }) {
  return <button className={`btn btn--${variant} ${className}`} {...props}>{icon}{children}</button>
}

export function GlassModal({ open, onClose, title, children, className = '' }: { open: boolean; onClose: () => void; title: string; children: ReactNode; className?: string }) {
  const dialog = useRef<HTMLDialogElement>(null)
  useEffect(() => {
    const node = dialog.current
    if (!node) return
    if (open && !node.open) node.showModal()
    if (!open && node.open) node.close()
  }, [open])
  return <dialog ref={dialog} className={`glass-dialog ${className}`} aria-label={title} onClose={onClose} onClick={event => { if (event.target === dialog.current) onClose() }}>
    <div className="glass-dialog__surface">
      <div className="glass-dialog__head"><h2>{title}</h2><button type="button" onClick={onClose} aria-label="Закрыть окно"><Icon name="close" size={20} /></button></div>
      {children}
    </div>
  </dialog>
}
export function Badge({ children, tone = 'neutral' }: { children: ReactNode; tone?: 'neutral' | 'success' | 'warning' | 'danger' | 'accent' }) {
  return <span className={`badge badge--${tone}`}>{children}</span>
}

export function ScreenHeading({ eyebrow, title, description, action }: { eyebrow?: string; title: string; description?: string; action?: ReactNode }) {
  return <div className="screen-heading"><div>{eyebrow && <p className="eyebrow">{eyebrow}</p>}<h1>{title}</h1>{description && <p className="screen-heading__description">{description}</p>}</div>{action && <div className="screen-heading__action">{action}</div>}</div>
}

export function EmptyState({ icon, title, description, action }: { icon: 'file' | 'grid' | 'shield' | 'clock'; title: string; description: string; action?: ReactNode }) {
  return <div className="empty-state"><div className="empty-state__icon"><Icon name={icon} size={25} /></div><h3>{title}</h3><p>{description}</p>{action && <div className="empty-state__action">{action}</div>}</div>
}

export function FileDrop({ label, hint, accept, multiple, disabled, onFiles }: { label: string; hint: string; accept: string; multiple?: boolean; disabled?: boolean; onFiles: (files: File[]) => void }) {
  const input = useRef<HTMLInputElement>(null)
  const [dragging, setDragging] = useState(false)
  function take(files: FileList | null) {
    if (!files?.length) return
    onFiles(Array.from(files))
    if (input.current) input.current.value = ''
  }
  function drop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault()
    setDragging(false)
    if (!disabled) take(event.dataTransfer.files)
  }
  function change(event: ChangeEvent<HTMLInputElement>) { take(event.target.files) }
  return <div className={`file-drop ${dragging ? 'file-drop--dragging' : ''} ${disabled ? 'file-drop--disabled' : ''}`} onDragOver={e => { e.preventDefault(); if (!disabled) setDragging(true) }} onDragLeave={() => setDragging(false)} onDrop={drop} role="button" tabIndex={disabled ? -1 : 0} onClick={() => !disabled && input.current?.click()} onKeyDown={e => { if (!disabled && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); input.current?.click() } }} aria-label={label}>
    <input ref={input} type="file" accept={accept} multiple={multiple} onChange={change} disabled={disabled} hidden />
    <span className="file-drop__icon"><Icon name="upload" size={22} /></span>
    <strong>{label}</strong><small>{hint}</small>
  </div>
}

export function SlidePreview({ url, index, className = '' }: { url?: string | null; index?: number; className?: string }) {
  return <div className={`slide-preview ${className}`}>
    {url ? <img src={resolveUrl(url)} alt={index !== undefined ? `Предпросмотр слайда ${index + 1}` : 'Предпросмотр слайда'} loading="lazy" /> : <div className="slide-preview__placeholder"><Icon name="layers" size={28} /><span>Предпросмотр пока недоступен</span></div>}
  </div>
}

export function ArtifactButtons({ version, onError }: { version: Version; onError: (message: string) => void }) {
  async function handle(artifact: Artifact) {
    try {
      const extension = artifact.kind.toLowerCase()
      await downloadArtifact(artifact, `lukas-variant-${version.variant_id || version.ordinal}.${extension}`)
    } catch (error) { onError(error instanceof Error ? error.message : 'Ошибка скачивания') }
  }
  // Slide PNGs are shown in the viewer; downloads are the deck files only.
  const files = version.artifacts?.filter(artifact => artifact.kind !== 'preview') || []
  if (!files.length) return <span className="muted small">Файлы появятся после экспорта</span>
  return <div className="artifact-actions">{files.map(a => <button type="button" key={a.id} onClick={() => handle(a)} className="artifact-link" title={`Скачать ${a.kind.toUpperCase()}`}><Icon name="download" size={15} />{a.kind.toUpperCase()}</button>)}</div>
}

export function formatDate(value?: string | null) {
  if (!value) return '—'
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) return value
  return new Intl.DateTimeFormat('ru-RU', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }).format(date)
}

export function ModeBadge({ version }: { version: Version }) {
  const metrics = version.plan?.metrics as { model_mode?: string } | undefined
  if (metrics?.model_mode === 'local_draft') return <Badge tone="warning">Локальный черновик · без ИИ</Badge>
  if (metrics?.model_mode === 'configured_api') return <Badge tone="accent">Модельный запуск</Badge>
  return null
}

export function qualityTone(status: string): 'success' | 'warning' | 'danger' | 'neutral' {
  if (['passed', 'approved', 'ready', 'completed'].includes(status)) return 'success'
  if (['failed', 'blocked', 'invalid'].includes(status)) return 'danger'
  if (['warning', 'warnings', 'completed_with_warnings', 'needs_review'].includes(status)) return 'warning'
  return 'neutral'
}
export function qualityLabel(status: string) {
  const labels: Record<string, string> = { passed: 'Проверен', approved: 'Одобрен', ready: 'Готов', completed: 'Готов', failed: 'Ошибка', blocked: 'Есть блокирующие ошибки', invalid: 'Некорректен', warning: 'Есть замечания', warnings: 'Есть замечания', completed_with_warnings: 'Есть замечания', needs_review: 'Нужна проверка', draft: 'Черновик' }
  return labels[status] || status || 'Без статуса'
}

// Every generation has three themes: the same facts, different layouts and amount of text.
const VARIANT_NAMES: Record<string, string> = { a: 'Базовая', b: 'Больше текста', c: 'Больше визуала' }

const VARIANT_HINTS: Record<string, string> = { a: 'Обычная подача по шаблону', b: 'Макеты с большим объёмом текста', c: 'Графики, крупные цифры и схемы' }

export function variantHint(version: { variant_id?: string | null }) {
  return VARIANT_HINTS[version.variant_id ?? ''] ?? ''
}

export function variantName(version: { variant_id?: string | null }, index: number) {
  return VARIANT_NAMES[version.variant_id ?? ''] ?? ['Базовая', 'Больше текста', 'Больше визуала'][index] ?? `Вариант ${index + 1}`
}

export function stageLabel(stage?: string | null) {
  if (!stage) return 'Подготовка задания'
  const labels: Record<string, string> = { ingest: 'Чтение материалов', template: 'Анализ шаблона', facts: 'Проверка фактов', plan: 'План презентации', planning: 'План презентации', slide_specs: 'Создание слайдов', slides: 'Создание слайдов', render: 'Сборка презентаций', audit: 'Проверка качества', export: 'Подготовка файлов', completed: 'Готово', ready: 'Готово', generating: 'Генерация', cancelling: 'Отмена', cancelled: 'Остановлено', repairing: 'Исправление', editing: 'Правка слайда', starting: 'Запуск', queued: 'В очереди', pending_enqueue: 'В очереди', running: 'Выполняется', analyzing: 'Анализ шаблона', interrupted: 'Прервано', failed: 'Ошибка' }
  // Service codes from the API are never shown as is; stages the engine reports are already in Russian.
  return labels[stage] || (/^[a-z_]+$/.test(stage) ? 'Выполняется' : stage)
}
