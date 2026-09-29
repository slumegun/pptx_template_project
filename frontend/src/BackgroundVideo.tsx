import { useEffect, useRef, useState } from 'react'

export default function BackgroundVideo({ className }: { className: string }) {
  const videoRef = useRef<HTMLVideoElement>(null)
  const [playBlocked, setPlayBlocked] = useState(false)

  useEffect(() => {
    const video = videoRef.current
    if (!video) return
    let disposed = false
    const start = async () => {
      try {
        video.muted = true
        await video.play()
        if (!disposed) setPlayBlocked(false)
      } catch {
        if (!disposed) setPlayBlocked(true)
      }
    }
    void start()
    const onVisible = () => { if (document.visibilityState === 'visible' && video.paused) void start() }
    document.addEventListener('visibilitychange', onVisible)
    video.addEventListener('canplay', start)
    const timeout = window.setTimeout(() => { if (video.paused) void start() }, 1200)
    return () => {
      disposed = true
      document.removeEventListener('visibilitychange', onVisible)
      video.removeEventListener('canplay', start)
      window.clearTimeout(timeout)
    }
  }, [])

  return <>
    <video ref={videoRef} className={className} src="/lukas-hero.mp4" poster="/lukas-poster.png" autoPlay muted loop playsInline preload="auto" aria-hidden="true" />
    {playBlocked && <button type="button" className="video-play" onClick={() => void videoRef.current?.play().then(() => setPlayBlocked(false))}>Включить видео</button>}
  </>
}
