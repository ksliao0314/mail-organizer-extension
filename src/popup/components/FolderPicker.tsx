import { useEffect, useMemo, useRef, useState } from 'react'
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
  const wrapRef = useRef<HTMLDivElement>(null)

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

  return (
    <div ref={wrapRef} className={cn('relative', className)}>
      <Input
        type="text"
        placeholder={placeholder}
        value={open ? query : value ?? ''}
        onChange={(e) => {
          setQuery(e.target.value)
          setOpen(true)
        }}
        onFocus={() => setOpen(true)}
        onKeyDown={(e) => {
          if (e.key === 'Escape') {
            setOpen(false)
            setQuery('')
          }
        }}
        className="font-mono text-xs"
      />
      {open && (
        <ul className="absolute top-full left-0 right-0 z-50 mt-1 max-h-56 overflow-y-auto rounded-md border border-border bg-popover shadow-md text-xs">
          {filtered.length === 0 ? (
            <li className="px-2 py-2 text-muted-foreground">沒有符合的資料夾</li>
          ) : (
            filtered.slice(0, 80).map((n) => {
              const pending = n.id.startsWith('pending:')
              return (
                <li
                  key={n.id}
                  role="option"
                  aria-selected={n.path === value}
                  onMouseDown={(e) => {
                    e.preventDefault()
                    onSelect(n)
                    setOpen(false)
                    setQuery('')
                  }}
                  className={cn(
                    'cursor-pointer px-2 py-1.5 font-mono hover:bg-accent hover:text-accent-foreground flex items-baseline gap-2',
                    n.path === value && 'bg-accent/60',
                    pending && 'bg-amber-50/60',
                  )}
                >
                  <span className="flex-1 truncate">{highlight(n.path)}</span>
                  {pending && (
                    <span className="text-[9px] text-amber-700 font-sans whitespace-nowrap">
                      待建立
                    </span>
                  )}
                </li>
              )
            })
          )}
          {filtered.length > 80 && (
            <li className="px-2 py-1 text-muted-foreground italic text-[10px]">…還有 {filtered.length - 80} 個未顯示，請繼續輸入縮小範圍</li>
          )}
        </ul>
      )}
    </div>
  )
}
