import { useEffect, useRef, useState } from 'react'
import type { KeyboardEvent } from 'react'
import { GlassModal, SlidePreview } from './ui'
import { Icon } from './icons'
import type { IconName } from './icons'

interface QuickAction { id: string; label: string; icon: IconName; instruction: string }

// The slide editor changes text and moves or resizes existing blocks, so every action stays within that.
const GROUPS: { title: string; actions: QuickAction[] }[] = [
  { title: 'Компоновка', actions: [
    { id: 'layout', label: 'Другая компоновка', icon: 'spark', instruction: 'Предложи другую компоновку: переставь блоки и измени их размеры, текст сохрани.' },
    { id: 'align', label: 'Выровнять блоки', icon: 'align', instruction: 'Выровняй блоки по общим краям и сделай отступы между ними одинаковыми.' },
    { id: 'visual', label: 'Сделать нагляднее', icon: 'eye', instruction: 'Сделай слайд нагляднее: выдели главную мысль, разбей текст на короткие пункты.' },
  ] },
  { title: 'Текст', actions: [
    { id: 'improve', label: 'Улучшить текст', icon: 'pen', instruction: 'Улучши текст: сделай формулировки яснее и живее, смысл и факты сохрани.' },
    { id: 'fix', label: 'Исправить ошибки', icon: 'check', instruction: 'Исправь орфографию, пунктуацию и грамматику, больше ничего не меняй.' },
    { id: 'shorter', label: 'Короче', icon: 'shorter', instruction: 'Сократи текст примерно вдвое, оставь главное.' },
    { id: 'longer', label: 'Подробнее', icon: 'longer', instruction: 'Раскрой мысль подробнее, но так, чтобы текст поместился в свои блоки.' },
    { id: 'simpler', label: 'Проще', icon: 'chat', instruction: 'Упрости язык: короткие предложения, без канцелярита и жаргона.' },
    { id: 'specific', label: 'Конкретнее', icon: 'target', instruction: 'Сделай формулировки конкретнее: факты и цифры из материалов вместо общих слов.' },
    { id: 'conclusion', label: 'Заголовок-вывод', icon: 'heading', instruction: 'Перепиши заголовок так, чтобы он формулировал вывод слайда, а не тему.' },
  ] },
]
const ACTIONS = GROUPS.flatMap(group => group.actions)

export default function SlideEditDialog({ open, slideNumber, variantLabel, previewUrl, onClose, onSubmit }: {
  open: boolean
  slideNumber: number
  variantLabel: string
  previewUrl?: string | null
  onClose: () => void
  onSubmit: (prompt: string) => Promise<unknown>
}) {
  const [text, setText] = useState('')
  const [picked, setPicked] = useState<string[]>([])
  const [busy, setBusy] = useState(false)
  const input = useRef<HTMLTextAreaElement>(null)

  useEffect(() => {
    if (!open) return
    setText(''); setPicked([])
    const focus = window.setTimeout(() => input.current?.focus(), 60)
    return () => window.clearTimeout(focus)
  }, [open, slideNumber])

  const prompt = [...ACTIONS.filter(action => picked.includes(action.id)).map(action => action.instruction), text.trim()]
    .filter(Boolean).join('\n')
  const ready = prompt.length >= 3 && !busy

  function toggle(id: string) {
    setPicked(previous => previous.includes(id) ? previous.filter(item => item !== id) : [...previous, id])
  }

  async function submit() {
    if (!ready) return
    setBusy(true)
    try { await onSubmit(prompt) } finally { setBusy(false) }
  }

  function keyDown(event: KeyboardEvent<HTMLTextAreaElement>) {
    if (event.key === 'Enter' && !event.shiftKey && !event.nativeEvent.isComposing) {
      event.preventDefault()
      void submit()
    }
  }

  return <GlassModal open={open} onClose={onClose} title={`Редактировать слайд ${slideNumber}`} className="edit-dialog slide-edit">
    <div className="slide-edit__context">
      <SlidePreview url={previewUrl} index={slideNumber - 1} className="slide-edit__thumb" />
      <p>Изменится только этот слайд в «{variantLabel}». Остальные слайды и варианты останутся как есть.</p>
    </div>
    <div className="slide-edit__composer">
      <textarea ref={input} value={text} onChange={event => setText(event.target.value)} onKeyDown={keyDown} rows={3} maxLength={1000}
        placeholder="Как изменить этот слайд?" aria-label="Что изменить на слайде" disabled={busy} />
      <div className="slide-edit__bar">
        <span>{busy ? 'Отправляем правку…' : picked.length ? `Выбрано действий: ${picked.length}` : 'Enter — применить, Shift+Enter — новая строка'}</span>
        <button type="button" className="slide-edit__send" onClick={() => void submit()} disabled={!ready} aria-label="Применить правки">
          {busy ? <span className="tiny-spinner" /> : <Icon name="arrow" size={18} />}
        </button>
      </div>
    </div>
    {GROUPS.map(group => <section key={group.title} className="slide-edit__group" aria-label={group.title}>
      <h3>{group.title}</h3>
      <div className="slide-edit__chips">{group.actions.map(action => <button type="button" key={action.id}
        className={picked.includes(action.id) ? 'slide-edit__chip slide-edit__chip--on' : 'slide-edit__chip'}
        aria-pressed={picked.includes(action.id)} onClick={() => toggle(action.id)} disabled={busy} title={action.instruction}>
        <Icon name={action.icon} size={15} />{action.label}
      </button>)}</div>
    </section>)}
  </GlassModal>
}
