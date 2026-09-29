import type { Run } from './types'
import { Button, stageLabel } from './ui'
import BackgroundVideo from './BackgroundVideo'

const completed = new Set(['completed', 'completed_with_warnings', 'awaiting_selection'])
const terminal = new Set([...completed, 'failed', 'interrupted', 'cancelled'])

export default function ProgressPage({ run, busy, onCancel, onRetry, onResults, onClose }: { run: Run | null; busy: string | null; onCancel: () => Promise<void>; onRetry: () => Promise<void>; onResults: () => void; onClose: () => void }) {
  if (!run) return <div className="page minimal-progress"><div className="minimal-progress__content"><h1>Пока нет задачи</h1><p>Создайте презентацию, чтобы увидеть ход работы.</p></div></div>

  const done = completed.has(run.status)
  const failed = run.status === 'failed' || run.status === 'interrupted'
  const active = !terminal.has(run.status)
  const editing = run.kind === 'repair' || run.kind === 'edit'
  const percent = typeof run.progress === 'number' && Number.isFinite(run.progress) ? Math.max(0, Math.min(100, run.progress)) : null
  const stage = done ? 'Готово' : failed ? 'Ошибка генерации' : run.status === 'cancelled' ? 'Остановлено' : run.status === 'queued' || run.status === 'pending_enqueue' ? 'Ожидаем запуск' : stageLabel(run.stage)
  const critic = done ? 'Критик завершил проверку композиции и содержания.' : failed ? 'Задачу можно запустить повторно.' : run.status === 'cancelled' ? 'Генерация остановлена.' : run.stage === 'audit' ? 'Критик проверяет композицию, читаемость и содержание.' : 'После сборки критик проверит композицию, читаемость и содержание.'

  return <div className={'page minimal-progress' + (active ? ' minimal-progress--active' : '')}>
    {active && <><BackgroundVideo className="progress-video" /><div className="progress-video__veil" /></>}
    <div className="minimal-progress__content">
      {run.status === 'cancelled' && <button type="button" className="minimal-progress__close" onClick={onClose} aria-label="Закрыть экран генерации">×</button>}
      <h1>{done ? (editing ? 'Правки готовы' : 'Презентации готовы') : failed ? 'Не удалось завершить' : run.status === 'cancelled' ? 'Генерация остановлена' : 'Создаём ваш шедевр'}</h1>
      <section className="minimal-progress__panel" aria-label="Ход генерации">
        <div className="minimal-progress__summary"><span>{stage}</span><strong>{percent === null ? '—' : Math.round(percent) + '%'}</strong></div>
        <div className="minimal-progress__track" role={percent === null ? undefined : 'progressbar'} aria-valuenow={percent ?? undefined} aria-valuemin={percent === null ? undefined : 0} aria-valuemax={percent === null ? undefined : 100}>
          <span style={{ width: (percent ?? 0) + '%' }} />
        </div>
        <p className="minimal-progress__critic">{critic}</p>
        {active && percent === null && <p className="minimal-progress__minor">Процент появится, когда сервер сообщит о ходе работы.</p>}
        {failed && run.error && <p className="minimal-progress__error">{run.error}</p>}
        {run.warnings && run.warnings.length > 0 && done && <p className="minimal-progress__minor">{run.warnings[0]}</p>}
        <div className="minimal-progress__actions">{done ? <Button onClick={onResults}>{editing ? 'Смотреть версию' : 'Смотреть варианты'}</Button> : failed ? <Button onClick={() => void onRetry()} disabled={Boolean(busy)}>Повторить</Button> : active ? <Button variant="secondary" onClick={() => void onCancel()} disabled={Boolean(busy)}>Остановить</Button> : null}</div>
      </section>
    </div>
  </div>
}
