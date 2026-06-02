import { describe, expect, it } from 'vitest'
import {
  encodeFolderName,
  findFolderByPath,
  joinFolderPath,
  PATH_SEP,
} from '@/shared/outlook-api'
import type { MailFolderNode } from '@/shared/types'

const FULLWIDTH = '／' // U+FF0F

describe('encodeFolderName', () => {
  it('returns name unchanged when no slash', () => {
    expect(encodeFolderName('客戶 A')).toBe('客戶 A')
  })

  it('replaces single slash with fullwidth solidus', () => {
    expect(encodeFolderName('客戶/A')).toBe(`客戶${FULLWIDTH}A`)
  })

  it('replaces multiple slashes', () => {
    expect(encodeFolderName('a/b/c')).toBe(`a${FULLWIDTH}b${FULLWIDTH}c`)
  })

  it('handles empty string', () => {
    expect(encodeFolderName('')).toBe('')
  })

  it('leaves existing fullwidth solidus untouched', () => {
    // user might already have a fullwidth slash in name (CJK input)
    expect(encodeFolderName(`客戶${FULLWIDTH}A`)).toBe(`客戶${FULLWIDTH}A`)
  })
})

describe('joinFolderPath', () => {
  it('returns just encoded name when parent is undefined', () => {
    expect(joinFolderPath(undefined, '案件')).toBe('案件')
  })

  it('returns just encoded name when parent is empty string', () => {
    expect(joinFolderPath('', '案件')).toBe('案件')
  })

  it('joins parent and child with PATH_SEP', () => {
    expect(joinFolderPath('案件', '客戶 A')).toBe('案件/客戶 A')
  })

  it('encodes slash in child name', () => {
    expect(joinFolderPath('案件', '客戶/A')).toBe(`案件${PATH_SEP}客戶${FULLWIDTH}A`)
  })

  it('uses canonical separator regardless of platform', () => {
    expect(PATH_SEP).toBe('/')
  })
})

describe('findFolderByPath with slashes in names', () => {
  function leaf(id: string, displayName: string, parentPath: string | undefined): MailFolderNode {
    return {
      id,
      displayName,
      parentFolderId: undefined,
      path: joinFolderPath(parentPath, displayName),
      children: [],
    }
  }

  it('disambiguates slash-in-name from real nesting', () => {
    // Folder "客戶/A" lives directly under "案件" (2 levels total).
    const slashInName = leaf('fid-1', '客戶/A', '案件')
    // Independently, a real 3-level structure: 案件 → 客戶 → A.
    const realA = leaf('fid-2', 'A', '案件/客戶')
    const realNested: MailFolderNode = {
      id: 'fid-客戶',
      displayName: '客戶',
      path: '案件/客戶',
      children: [realA],
      parentFolderId: 'fid-root',
    }
    const root: MailFolderNode = {
      id: 'fid-root',
      displayName: '案件',
      path: '案件',
      children: [slashInName, realNested],
      parentFolderId: undefined,
    }
    const tree = [root]

    // The slash-in-name node has path with U+FF0F; the 3-level has U+002F.
    expect(slashInName.path).toBe(`案件${PATH_SEP}客戶${FULLWIDTH}A`)
    expect(realA.path).toBe('案件/客戶/A')

    // Each path resolves to its own node — no collision.
    expect(findFolderByPath(tree, slashInName.path)).toBe(slashInName)
    expect(findFolderByPath(tree, realA.path)).toBe(realA)
  })

  it('returns undefined for a path that nothing in tree matches', () => {
    const node = leaf('x', 'X', undefined)
    expect(findFolderByPath([node], 'does/not/exist')).toBeUndefined()
  })
})
