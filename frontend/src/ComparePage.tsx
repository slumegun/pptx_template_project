import { useEffect, useRef, useState } from 'react'
import type { Version } from './types'
import { ArtifactButtons, Badge, Button, EmptyState, GlassModal, ModeBadge, ScreenHeading, SlidePreview, qualityLabel, qualityTone, variantHint, variantName } from './ui'
import { Icon } from './icons'
import SlideEditDialog from './SlideEditDialog'

export default function ComparePage({ versions, onSelect, onAudit, onEdit, onNew, onError }: { versions: Version[]; selectedId: string | null; onSelect: (id: string) => void; onAudit: (id: string) => void; onEdit: (id: string, slideIndex: number, prompt: string) => Promise<boolean>; onNew: () => void; onError: (message: string) => void }) {
  const [viewerId, setViewerId] = useState<string | null>(null)
  const [slideIndex, setSlideIndex] = useState(0)
  const [editOpen, setEditOpen] = useState(false)
  const viewer = versions.find(version => version.id === viewerId) || null
  const isRepair = versions.length === 1 && Boolean(versions[0]?.parent_version_id)
  const slideCount = viewer?.preview_urls?.length || 0
  const safeSlideIndex = Math.min(slideIndex, Math.max(0, slideCount - 1))
  const rail = useRef<HTMLDivElement>(null)

  useEffect(() => {
    rail.current?.children[safeSlideIndex]?.scrollIntoView({ block: 'nearest', inline: 'nearest' })
  }, [safeSlideIndex, viewerId])

  useEffect(() => {
    if (!viewerId || editOpen) return
    function step(event: KeyboardEvent) {
      if (event.key === 'ArrowRight') setSlideIndex(index => Math.min(index + 1, Math.max(0, slideCount - 1)))
      if (event.key === 'ArrowLeft') setSlideIndex(index => Math.max(index - 1, 0))
    }
    window.addEventListener('keydown', step)
    return () => window.removeEventListener('keydown', step)
  }, [viewerId, editOpen, slideCount])

  function openViewer(version: Version) {
    onSelect(version.id)
    setViewerId(version.id)
    setSlideIndex(0)
  }

  function openEdit() {
    setEditOpen(true)
  }

  async function applyEdit(prompt: string) {
    if (!viewer) return
    const started = await onEdit(viewer.id, safeSlideIndex + 1, prompt)
    if (started) { setEditOpen(false); setViewerId(null) }
  }

  return <div className="page page--compare">
    <ScreenHeading title={isRepair ? 'Исправленная версия' : 'Три взгляда на одну идею'} description={isRepair ? 'Проверьте исправления и скачайте новую версию.' : 'У всех вариантов одинаковое число слайдов. Выберите подачу, которая вам ближе.'} action={<Button variant="secondary" onClick={onNew} icon={<Icon name="plus" size={17} />}>Новая презентация</Button>} />
    {versions.length === 0 ? <EmptyState icon="grid" title="Варианты ещё не созданы" description="После генерации здесь появятся презентации для сравнения." action={<Button onClick={onNew}>Начать создание</Button>} /> : <>
      {versions.length < 3 && !isRepair && <div className="notice notice--warning"><Icon name="info" size={18} />Пока доступно {versions.length} из трёх вариантов. Проверьте статус задачи.</div>}
      <div className="variant-grid">{versions.map((version, index) => <article key={version.id} className="variant-card">
        <button type="button" className="variant-card__preview-button" onClick={() => openViewer(version)} aria-label={`Открыть вариант ${index + 1}`}><SlidePreview url={version.preview_urls?.[0]} /><span className="variant-card__number">0{index + 1}</span></button>
        <div className="variant-card__content"><div className="variant-card__top"><h2>{isRepair ? 'Исправленная версия' : variantName(version, index)}</h2><Badge tone={qualityTone(version.quality_status)}>{qualityLabel(version.quality_status)}</Badge></div><div className="variant-card__mode"><ModeBadge version={version} /></div><p>{!isRepair && variantHint(version) ? `${variantHint(version)} · ` : ''}{version.preview_urls?.length ? `${version.preview_urls.length} слайдов` : 'Презентация подготовлена'}</p><div className="variant-card__actions"><Button onClick={() => openViewer(version)} icon={<Icon name="eye" size={16} />}>Открыть</Button><button type="button" className="text-action" onClick={() => onAudit(version.id)}>Аудит <Icon name="arrow" size={15} /></button></div><ArtifactButtons version={version} onError={onError} /></div>
      </article>)}</div>
    </>}
    <GlassModal open={Boolean(viewer)} onClose={() => { setViewerId(null); setEditOpen(false) }} title={viewer ? (isRepair ? 'Исправленная версия' : variantName(viewer, versions.indexOf(viewer))) : 'Презентация'} className={editOpen ? 'deck-dialog deck-dialog--editing' : 'deck-dialog'}>
      {viewer && <div className="modal-deck">
        <div className="modal-deck__stage"><button type="button" className="modal-deck__slide" onClick={openEdit} disabled={!slideCount} title="Нажмите на слайд, чтобы описать правки" aria-label={`Предложить правки для слайда ${safeSlideIndex + 1}`}><SlidePreview url={viewer.preview_urls?.[safeSlideIndex]} index={safeSlideIndex} /></button></div>
        <p className="modal-deck__hint">{slideCount ? `Слайд ${safeSlideIndex + 1} из ${slideCount} · нажмите на слайд, чтобы описать правки` : 'Предпросмотр недоступен'}</p>
        <div className="modal-deck__rail" ref={rail} onWheel={event => { if (rail.current && Math.abs(event.deltaY) > Math.abs(event.deltaX)) rail.current.scrollLeft += event.deltaY }}>{viewer.preview_urls?.map((url, index) => <button type="button" key={index} className={safeSlideIndex === index ? 'modal-deck__thumb active' : 'modal-deck__thumb'} onClick={() => setSlideIndex(index)} aria-label={`Показать слайд ${index + 1}`} aria-current={safeSlideIndex === index}><SlidePreview url={url} index={index} /><span>{String(index + 1).padStart(2, '0')}</span></button>)}</div>
      </div>}
    </GlassModal>
    <SlideEditDialog open={editOpen} slideNumber={safeSlideIndex + 1} previewUrl={viewer?.preview_urls?.[safeSlideIndex]}
      variantLabel={viewer ? (isRepair ? 'Исправленная версия' : variantName(viewer, versions.indexOf(viewer))) : ''}
      onClose={() => setEditOpen(false)} onSubmit={applyEdit} />
  </div>
}
