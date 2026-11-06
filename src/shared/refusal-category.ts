// Refusal categories in user-facing text (2026-11-06).
//
// stop_details.category is an open set of API ids (`cyber`, `bio`,
// `frontier_llm`, `reasoning_extraction`, `general_harms`, … or null). Plan
// rows, banners and the refused-email ledger used to show those raw ids. One
// neutral Traditional Chinese label per category, shared by the service
// worker's texts and the popup. The raw id stays in the error log context for
// diagnosis. Dependency-free: the popup imports it.

const LABELS: Readonly<Record<string, string>> = {
  cyber: '資訊安全',
  bio: '生物安全',
  frontier_llm: 'AI 模型開發',
  reasoning_extraction: '推理過程擷取',
  general_harms: '一般使用政策',
}

/** Neutral label for a refusal category; null / empty = 「未提供」, unknown ids = 「其他」. */
export function refusalCategoryLabel(category: string | null | undefined): string {
  if (!category) return '未提供'
  return LABELS[category] ?? '其他'
}

/** Distinct labels of several categories, in first-seen order, joined with 「、」. */
export function refusalCategoriesText(categories: Iterable<string | null | undefined>): string {
  const labels: string[] = []
  for (const c of categories) {
    const label = refusalCategoryLabel(c)
    if (!labels.includes(label)) labels.push(label)
  }
  return labels.length > 0 ? labels.join('、') : refusalCategoryLabel(null)
}
