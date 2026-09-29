import { useState } from 'react'
import type { FormEvent } from 'react'
import { Button } from './ui'
import { Icon } from './icons'
import type { AuthMode } from './HomePage'

export default function AuthPage({ mode, onMode, onBack, onSubmit, busy, error }: {
  mode: AuthMode
  onMode: (mode: AuthMode) => void
  onBack: () => void
  onSubmit: (mode: AuthMode, email: string, password: string) => Promise<void>
  busy: boolean
  error: string | null
}) {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [showPassword, setShowPassword] = useState(false)
  function submit(event: FormEvent) {
    event.preventDefault()
    void onSubmit(mode, email, password)
  }
  return <div className="auth-page">
    <div className="auth-page__image" />
    <div className="auth-page__shade" />
    <header className="auth-page__header"><button type="button" className="auth-page__back" onClick={onBack} aria-label="На главную"><Icon name="arrow" size={21} /></button><button className="home__wordmark" type="button" onClick={onBack}>ЛУКАС<span>.</span></button></header>
    <main className="auth-page__main">
      <section className="auth-panel">
        <div className="auth-panel__eyebrow">ВАШЕ ПРОСТРАНСТВО ИДЕЙ</div>
        <h1>{mode === 'login' ? 'С возвращением' : 'Создайте аккаунт'}</h1>
        <p>{mode === 'login' ? 'Войдите, чтобы продолжить работу над презентациями.' : 'Зарегистрируйтесь, чтобы сохранять шаблоны и презентации.'}</p>
        <div className="auth-tabs" role="tablist" aria-label="Вход или регистрация"><button type="button" role="tab" aria-selected={mode === 'login'} className={mode === 'login' ? 'active' : ''} onClick={() => onMode('login')}>Войти</button><button type="button" role="tab" aria-selected={mode === 'register'} className={mode === 'register' ? 'active' : ''} onClick={() => onMode('register')}>Зарегистрироваться</button></div>
        <form className="auth-form" onSubmit={submit}>
          <label>Электронная почта<input type="email" autoComplete="email" value={email} onChange={event => setEmail(event.target.value)} placeholder="you@example.com" required /></label>
          <label>Пароль<span className="auth-password"><input type={showPassword ? 'text' : 'password'} autoComplete={mode === 'login' ? 'current-password' : 'new-password'} minLength={8} maxLength={128} value={password} onChange={event => setPassword(event.target.value)} placeholder="Не менее 8 символов" required /><button type="button" onClick={() => setShowPassword(value => !value)} aria-label={showPassword ? 'Скрыть пароль' : 'Показать пароль'}>{showPassword ? 'Скрыть' : 'Показать'}</button></span></label>
          {error && <div className="auth-panel__error" role="alert">{error}</div>}
          <Button type="submit" disabled={busy}>{busy ? 'Подождите…' : mode === 'login' ? 'Войти' : 'Зарегистрироваться'} <Icon name="arrow" size={17} /></Button>
        </form>
        <p className="auth-panel__switch">{mode === 'login' ? 'Ещё нет аккаунта?' : 'Уже есть аккаунт?'} <button type="button" onClick={() => onMode(mode === 'login' ? 'register' : 'login')}>{mode === 'login' ? 'Зарегистрироваться' : 'Войти'}</button></p>
      </section>
    </main>
    <footer className="auth-page__footer">ЛУКАС<span>.</span> · ИДЕИ ОБРЕТАЮТ ФОРМУ</footer>
  </div>
}
