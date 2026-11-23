/* Reference video for a concept: embedded YouTube player that skips unavailable videos automatically. */
import { useEffect, useRef, useState } from 'react'
import { CONCEPTS } from '../../mentor/concepts'
import type { MentorStore } from '../../mentor/store'
import type { ConceptId } from '../../mentor/types'
import { Icon } from '../icons'

interface YTPlayer {
  destroy(): void
}
interface YTNamespace {
  Player: new (el: HTMLElement, opts: Record<string, unknown>) => YTPlayer
}
declare global {
  interface Window {
    YT?: YTNamespace
    onYouTubeIframeAPIReady?: () => void
  }
}

let apiPromise: Promise<YTNamespace> | null = null
function loadYouTubeApi(): Promise<YTNamespace> {
  if (window.YT?.Player) return Promise.resolve(window.YT)
  if (apiPromise) return apiPromise
  apiPromise = new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('timeout')), 9000)
    const prev = window.onYouTubeIframeAPIReady
    window.onYouTubeIframeAPIReady = () => {
      prev?.()
      clearTimeout(timer)
      if (window.YT) resolve(window.YT)
    }
    const s = document.createElement('script')
    s.src = 'https://www.youtube.com/iframe_api'
    s.async = true
    s.onerror = () => {
      clearTimeout(timer)
      apiPromise = null
      reject(new Error('blocked'))
    }
    document.head.appendChild(s)
  })
  return apiPromise
}

export function VideoPlayer({ concept, store, compact }: { concept: ConceptId; store: MentorStore; compact?: boolean }) {
  const def = CONCEPTS[concept]
  const [bad, setBad] = useState(() => store.badVideos())
  const candidates = def.videos.filter((v) => !bad.has(v.id))
  const video = candidates[0]
  const host = useRef<HTMLDivElement>(null)
  const [state, setState] = useState<'idle' | 'loading' | 'ready' | 'offline'>('idle')
  const [play, setPlay] = useState(false)
  const search = `https://www.youtube.com/results?search_query=${encodeURIComponent(`Excel ${def.name} tutorial for beginners`)}`

  useEffect(() => {
    if (!play || !video || !host.current) return
    let player: YTPlayer | null = null
    let live = true
    setState('loading')
    const mount = document.createElement('div')
    host.current.innerHTML = ''
    host.current.appendChild(mount)
    loadYouTubeApi()
      .then((YT) => {
        if (!live) return
        player = new YT.Player(mount, {
          videoId: video.id,
          host: 'https://www.youtube-nocookie.com',
          width: '100%',
          height: '100%',
          playerVars: { rel: 0, modestbranding: 1, playsinline: 1 },
          events: {
            onReady: () => live && setState('ready'),
            // 2 invalid id, 5 html5 error, 100 removed/private, 101/150 embedding not allowed
            onError: () => {
              if (!live) return
              store.markVideoBad(video.id)
              setBad(store.badVideos())
            },
          },
        })
      })
      .catch(() => live && setState('offline'))
    return () => {
      live = false
      try {
        player?.destroy()
      } catch {
        /* ignore */
      }
    }
  }, [play, video, store])

  return (
    <div className="overflow-hidden rounded-lg border border-[var(--border)] bg-[var(--panel)]" data-testid="video-card">
      <div className="flex items-center gap-2 px-2.5 py-1.5 text-[12px]">
        <Icon name="eye" size={14} className="text-[var(--accent)]" />
        <span className="min-w-0 flex-1 truncate font-medium" title={video?.title}>{video ? video.title : `Videos about ${def.name}`}</span>
        {video && <span className="shrink-0 text-[11px] text-[var(--muted)]">{video.source}</span>}
      </div>
      {video && state !== 'offline' ? (
        play ? (
          <div className={compact ? 'aspect-video w-full bg-black' : 'aspect-video w-full bg-black'} ref={host}>
            <div className="flex h-full items-center justify-center text-[12px] text-white/70">Loading video…</div>
          </div>
        ) : (
          <button type="button" onClick={() => setPlay(true)} className="relative block aspect-video w-full bg-[#111] text-white" aria-label={`Play video: ${video.title}`}>
            <img src={`https://i.ytimg.com/vi/${video.id}/mqdefault.jpg`} alt="" className="h-full w-full object-cover opacity-80" onError={(e) => ((e.target as HTMLImageElement).style.display = 'none')} />
            <span className="absolute inset-0 flex items-center justify-center">
              <span className="flex h-11 w-16 items-center justify-center rounded-xl bg-[#e00]/90 text-[18px]">▶</span>
            </span>
          </button>
        )
      ) : (
        <div className="px-2.5 pb-2 text-[12px] text-[var(--muted)]">
          {state === 'offline' ? 'The video player could not load here (offline or blocked).' : 'No approved video is available for this concept right now.'}
        </div>
      )}
      <div className="flex flex-wrap gap-x-3 gap-y-1 border-t border-[var(--border)] px-2.5 py-1.5 text-[11.5px]">
        {video && (
          <a className="text-[var(--accent)] hover:underline" href={`https://www.youtube.com/watch?v=${video.id}`} target="_blank" rel="noopener noreferrer">
            Open on YouTube
          </a>
        )}
        <a className="text-[var(--accent)] hover:underline" href={search} target="_blank" rel="noopener noreferrer">
          Find another video
        </a>
        {def.article && (
          <a className="text-[var(--accent)] hover:underline" href={def.article.url} target="_blank" rel="noopener noreferrer">
            {def.article.title}
          </a>
        )}
      </div>
    </div>
  )
}
