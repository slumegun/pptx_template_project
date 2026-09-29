import { Button } from './ui'
import { Icon } from './icons'
import BackgroundVideo from './BackgroundVideo'

export type AuthMode = 'login' | 'register'

export default function HomePage({ onAuth }: { onAuth: (mode: AuthMode) => void }) {
  return <div className="home">
    <BackgroundVideo className="home__video" />
    <div className="home__veil" />
    <header className="home__header">
      <button className="home__wordmark" type="button" onClick={() => onAuth('login')} aria-label="Лукас — вход">ЛУКАС<span>.</span></button>
      <nav aria-label="Навигация"><button type="button" onClick={() => onAuth('login')}>Создать презентацию</button><button type="button" onClick={() => onAuth('login')}>Войти</button><Button onClick={() => onAuth('register')}>Регистрация <Icon name="arrow" size={16} /></Button></nav>
    </header>
    <main className="home__main">
      <div className="home__frame">
        <div className="home__intro"><span className="home__eyebrow"><span /> ПРОСТРАНСТВО ИДЕЙ И ФОРМЫ</span><h1>Лукас<span className="home__period">.</span></h1><blockquote>«Искусство — это хаос,<br /> взятый в рамку»<cite>— Марсель Пруст</cite></blockquote></div>
        <div className="home__bottom"><div className="home__lead"><span className="home__index">01 / СОЗДАНИЕ</span><p>Ваши идеи становятся<br /> презентациями.</p></div><button type="button" className="home__start" onClick={() => onAuth('login')}>Начать создание <span><Icon name="arrow" size={21} /></span></button></div>
      </div>
    </main>
    <footer className="home__footer"><span>ЛУКАС<span className="home__period">.</span> © 2026</span><span>ПРЕЗЕНТАЦИИ, В КОТОРЫХ ЕСТЬ ИДЕЯ</span></footer>
  </div>
}
