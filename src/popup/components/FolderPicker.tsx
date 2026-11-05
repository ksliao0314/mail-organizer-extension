import { useEffect, useId, useMemo, useRef, useState } from 'react'
import { Input } from '@/components/ui/input'
import { cn } from '@/lib/utils'
import type { MailFolderNode } from '@/shared/types'
import { flattenFolderTree } from '@/shared/outlook-api'

export type FolderPickerProps = {
  tree: MailFolderNode[]
  value?: string
  excludePrefixes?: string[]
  onSelect: (node: MailFolderNode) => void
  placeholder?: string
  className?: string
}

// Rendered option cap — a 500-folder tree would otherwise mount 500 rows on
// every keystroke; the hint row tells the user to keep typing.
const MAX_OPTIONS = 80

function pathExcluded(path: string, excludes: string[]): boolean {
  return excludes.some((p) => path === p || path.startsWith(p + '/'))
}

export function FolderPicker({
  tree,
  value,
  excludePrefixes = [],
  onSelect,
  placeholder = '搜尋資料夾…',
  className,
}: FolderPickerProps) {
  const [query, setQuery] = useState('')
  const [open, setOpen] = useState(false)
  // Keyboard cursor inside the open list (2026-11-05). -1 = none yet.
  const [activeIndex, setActiveIndex] = useState(-1)
  const wrapRef = useRef<HTMLDivElement>(null)
  const listRef = useRef<HTMLUListElement>(null)
  const baseId = useId()
  const listboxId = `${baseId}-listbox`
  const optionId = (i: number) => `${baseId}-opt-${i}`

  const flat = useMemo(() => {
    return flattenFolderTree(tree)
      .filter((n) => !pathExcluded(n.path, excludePrefixes))
      .sort((a, b) => a.path.localeCompare(b.path, 'zh-Hant'))
  }, [tree, excludePrefixes])

  const filtered = useMemo(() => {
    if (!query.trim()) return flat
    const q = query.toLowerCase()
    return flat.filter((n) => n.path.toLowerCase().includes(q))
  }, [flat, query])

  const visible = useMemo(() => filtered.slice(0, MAX_OPTIONS), [filtered])

  // A new query (or reopening) restarts the cursor at the top match so
  // Enter picks what the user most likely means.
  useEffect(() => {
    setActiveIndex(open && visible.length > 0 ? 0 : -1)
  }, [open, query, visible.length])

  // Keep the active option in view while arrowing through a long list.
  // (scrollIntoView is optional-called: jsdom doesn't implement it.)
  useEffect(() => {
    if (!open || activeIndex < 0) return
    const el = listRef.current?.querySelector<HTMLElement>(`[data-index="${activeIndex}"]`)
    el?.scrollIntoView?.({ block: 'nearest' })
  }, [open, activeIndex])

  useEffect(() => {
    function onClickOutside(e: MouseEvent) {
      if (!wrapRef.current) return
      if (!wrapRef.current.contains(e.target as Node)) {
        setOpen(false)
        setQuery('')
      }
    }
    if (open) document.addEventListener('mousedown', onClickOutside)
    return () => document.removeEventListener('mousedown', onClickOutside)
  }, [open])

  function choose(n: MailFolderNode) {
    onSelect(n)
    setOpen(false)
    setQuery('')
  }

  function highlight(text: string): React.ReactNode {
    if (!query.trim()) return text
    const q = query.toLowerCase()
    const lc = text.toLowerCase()
    const i = lc.indexOf(q)
    if (i < 0) return text
    return (
      <>
        {text.slice(0, i)}
        <mark className="bg-amber-200/60 text-foreground rounded-sm px-0.5">
          {text.slice(i, i + query.length)}
        </mark>
        {text.slice(i + query.length)}
      </>
    )
  }

  // Combobox keyboard contract (2026-11-05): ↑/↓ move the active option,
  // Enter picks it, Esc closes ONLY the dropdown. Handled keys stop
  // propagation so the document-level Esc handlers (rule drawer, OWA panel,
  // options modal) don't also fire — Esc on an open dropdown used to close
  // the whole rule drawer and discard the edit.
  function onKeyDown(e: React.KeyboardEvent<HTMLInputElement>) {
    if (e.nativeEvent.isComposing || e.keyCode === 229) return // IME
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault()
      e.stopPropagation()
      if (!open) {
        setOpen(true)
        return
      }
      if (visible.length === 0) return
      const delta = e.key === 'ArrowDown' ? 1 : -1
      setActiveIndex((i) => {
        if (i < 0) return delta > 0 ? 0 : visible.length - 1
        return (i + delta + visible.length) % visible.length
      })
      return
    }
    if (e.key === 'Enter') {
      if (!open) return
      e.preventDefault()
      e.stopPropagation()
      const n = activeIndex >= 0 ? visible[activeIndex] : undefined
      if (n) choose(n)
      return
    }
    if (e.key === 'Escape') {
      if (!open) return // closed: Esc belongs to the next layer out
      e.preventDefault()
      e.stopPropagation()
      setOpen(false)
      setQuery('')
      return
    }
    if (e.key === 'Tab' && open) {
      // Leaving the field by keyboard closes the list (outside-mousedown
      // only covers the pointer).
      setOpen(false)
      setQuery('')
    }
  }

  const activeOptionId = open && activeIndex >= 0 && activeIndex < visible.length
    ? optionId(activeIndex)
    : undefined

  return (
    <div ref={wrapRef} className={cn('relative', className)}>
      <Input
        type="text"
        role="combobox"
        aria-expanded={open}
        aria-controls={open ? listboxId : undefined}
        aria-autocomplete="list"
        aria-activedescendant={activeOptionId}
        placeholder={placeholder}
        value={open ? query : value ?? ''}
        onChange={(e) => {
          setQuery(e.target.value)
          setOpen(true)
        }}
        onFocus={() => setOpen(true)}
        onKeyDown={onKeyDown}
        className="font-mono text-xs"
      />
      {open && (
        <ul
          ref={listRef}
          id={listboxId}
          role="listbox"
          aria-label="資料夾"
          className="absolute top-full left-0 right-0 z-50 mt-1 max-h-56 overflow-y-auto rounded-md border border-border bg-popover shadow-md text-xs"
        >
          {filtered.length === 0 ? (
            <li className="px-2 py-2 text-muted-foreground">沒有符合的資料夾</li>
          ) : (
            visible.map((n, i) => {
              const pending = n.id.startsWith('pending:')
              const active = i === activeIndex
              return (
                <li
                  key={n.id}
                  id={optionId(i)}
                  data-index={i}
                  role="option"
                  aria-selected={active}
                  onMouseDown={(e) => {
                    e.preventDefault()
                    choose(n)
                  }}
                  onMouseEnter={() => setActiveIndex(i)}
                  className={cn(
                    'cursor-pointer px-2 py-1.5 font-mono hover:bg-accent hover:text-accent-foreground flex items-baseline gap-2',
                    n.path === value && 'bg-accent/60',
                    pending && 'bg-amber-50/60',
                    active && 'bg-accent text-accent-foreground',
                  )}
                >
                  <span className="flex-1 truncate">{highlight(n.path)}</span>
                  {n.path === value && (
                    <span className="text-[9px] text-muted-foreground font-sans whitespace-nowrap">
                      目前
                    </span>
                  )}
                  {pending && (
                    <span className="text-[9px] text-amber-700 font-sans whitespace-nowrap">
                      待建立
                    </span>
                  )}
                </li>
              )
            })
          )}
          {filtered.length > MAX_OPTIONS && (
            <li className="px-2 py-1 text-muted-foreground italic text-[10px]">…還有 {filtered.length - MAX_OPTIONS} 個未顯示，請繼續輸入縮小範圍</li>
          )}
        </ul>
      )}
    </div>
  )
}
