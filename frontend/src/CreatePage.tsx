import { useEffect, useRef, useState } from 'react'
import type { Project, Source } from './types'
import { Button, FileDrop, GlassModal, stageLabel } from './ui'

export interface CreatePageProps {
  project: Project | null
  template: Source | null
  templates: Source[]
  onSelectTemplate: (source: Source) => void
  onDeleteTemplate: (source: Source) => Promise<void>
  onExportTemplate: (source: Source) => Promise<void>
  onImportTemplate: (file: File) => Promise<void>
  brief: string
  slideCount: number
  busy: string | null
  onTemplateFile: (file: File) => Promise<void>
  onBriefChange: (value: string) => void
  onSlideCountChange: (value: number) => void
  onRetryPreparation: () => Promise<void>
  onGenerate: () => Promise<void>
}

export default function CreatePage(props: CreatePageProps) {
  const [promptOpen, setPromptOpen] = useState(false)
  const packageInput = useRef<HTMLInputElement>(null)
  useEffect(() => {
    if (props.slideCount < 1 || props.slideCount > 50) props.onSlideCountChange(Math.max(1, Math.min(50, props.slideCount)))
  }, [props.slideCount, props.onSlideCountChange])
  const slideCount = Math.max(1, Math.min(50, props.slideCount))
  const status = props.template?.preparation?.status || 'not_started'
  const ready = status === 'ready'
  const canGenerate = Boolean(props.project && props.template && ready && props.brief.trim() && props.slideCount >= 1 && props.slideCount <= 50 && !props.busy)

  return <div className="page minimal-create">
    <header className="minimal-create__heading">
      <h1>Создавай. Твори. Меняй<span className="minimal-create__dot">.</span></h1>
      <p>Давайте воплотим вашу идею в презентацию.</p>
    </header>

    {!props.project ? <p className="minimal-create__loading">Загружаем шаблоны…</p> : <div className="minimal-create__layout">
      <section className="minimal-create__section" aria-labelledby="templates-title">
        <div className="minimal-create__section-head"><h2 id="templates-title">Шаблон</h2>
          <button type="button" className="minimal-import" onClick={() => packageInput.current?.click()} disabled={Boolean(props.busy)} title="Добавить шаблон из пакета .zip, скачанного кнопкой «Скачать»: без повторного анализа">{props.busy === 'import' ? 'Загружаем…' : 'Загрузить готовый шаблон'}</button>
          <input ref={packageInput} type="file" accept=".zip,application/zip" hidden onChange={event => { const file = event.target.files?.[0]; event.target.value = ''; if (file) void props.onImportTemplate(file) }} />
        </div>
        {props.templates.length > 0 && <div className="minimal-templates" role="group" aria-label="Загруженные шаблоны">
          {props.templates.map(source => <div key={source.id} className={props.template?.id === source.id ? 'minimal-template minimal-template--selected' : 'minimal-template'}>
            <button type="button" className="minimal-template__select" onClick={() => props.onSelectTemplate(source)} disabled={Boolean(props.busy)}>
              <span className="minimal-template__name" title={source.filename}>{source.filename}</span>
              <span className="minimal-template__state">{source.preparation?.status === 'ready' ? 'Готов' : source.preparation?.status === 'failed' ? 'Ошибка' : 'Подготовка'}</span>
            </button>
            {source.preparation?.status === 'ready' && <button type="button" className="minimal-template__action" onClick={() => void props.onExportTemplate(source)} disabled={Boolean(props.busy)} aria-label={'Скачать пакет шаблона ' + source.filename} title="Скачать шаблон с анализом и шрифтами одним .zip">Скачать</button>}
            <button type="button" className="minimal-template__delete" onClick={() => void props.onDeleteTemplate(source)} disabled={Boolean(props.busy)} aria-label={'Удалить шаблон ' + source.filename}>Удалить</button>
          </div>)}
        </div>}
        <FileDrop label={props.template ? 'Загрузить другой шаблон' : 'Выбрать шаблон'} hint="Перетащите файл PPTX сюда или нажмите" accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation" disabled={Boolean(props.busy)} onFiles={files => files[0] && void props.onTemplateFile(files[0])} />
        {props.template && !ready && status !== 'failed' && <p className="minimal-create__status"><span className="tiny-spinner" /> {stageLabel(props.template.preparation?.stage) || 'Подготавливаем шаблон'}</p>}
        {status === 'failed' && <div className="minimal-create__error"><p>{props.template?.preparation?.error || 'Не удалось подготовить шаблон.'}</p><button type="button" onClick={() => void props.onRetryPreparation()} disabled={Boolean(props.busy)}>Повторить подготовку</button></div>}
      </section>

      <section className="minimal-create__section minimal-create__settings" aria-labelledby="settings-title">
        <div className="minimal-create__section-head"><h2 id="settings-title">Настройки</h2></div>
        <label className="minimal-count" htmlFor="slide-count-range"><span>Слайдов в каждом варианте</span><output htmlFor="slide-count-range">{slideCount}</output></label>
        <div className="minimal-range"><input id="slide-count-range" type="range" min="1" max="50" value={slideCount} onChange={event => props.onSlideCountChange(Number(event.target.value))} /><div><span>1</span><span>50</span></div></div>
        <button type="button" className="minimal-prompt" onClick={() => setPromptOpen(true)}><span><strong>Задание</strong><small>{props.brief.trim() || 'Опишите тему и нужное содержание'}</small></span><span aria-hidden="true">↗</span></button>
        <div className="minimal-create__action"><Button onClick={() => void props.onGenerate()} disabled={!canGenerate}>{props.busy === 'generate' ? 'Запускаем…' : 'Создать 3 варианта'}</Button>{!canGenerate && !props.busy && <p role="status">{!props.template ? 'Сначала выберите шаблон.' : !ready ? 'Дождитесь подготовки шаблона.' : 'Добавьте задание.'}</p>}</div>
      </section>
    </div>}

    <GlassModal open={promptOpen} onClose={() => setPromptOpen(false)} title="Задание" className="prompt-dialog">
      <label className="field-label" htmlFor="brief">Что должно быть в презентации?</label>
      <textarea id="brief" className="brief-input" value={props.brief} onChange={event => props.onBriefChange(event.target.value)} placeholder="Тема, аудитория, ключевые мысли…" rows={8} />
      <div className="glass-dialog__actions"><Button variant="secondary" onClick={() => setPromptOpen(false)}>Отмена</Button><Button onClick={() => setPromptOpen(false)} disabled={!props.brief.trim()}>Сохранить</Button></div>
    </GlassModal>
  </div>
}
