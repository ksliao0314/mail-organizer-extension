// Messages API SSE fixtures for classify tests (2026-11-05).
//
// The classifier streams (`stream: true`), so the API answers with
// Server-Sent Events. These helpers build Responses whose body arrives in
// small chunks — cutting events, lines and multi-byte CJK characters apart —
// exactly as a real network read does.

export type SseEvent = { event: string; data: unknown }

export function sseText(events: SseEvent[], lineEnding = '\n'): string {
  return events
    .map((e) => `event: ${e.event}${lineEnding}data: ${JSON.stringify(e.data)}${lineEnding}${lineEnding}`)
    .join('')
}

/** A streamed Response; `chunkBytes` splits the UTF-8 body into tiny reads. */
export function sseResponse(body: string, opts: { chunkBytes?: number; status?: number } = {}): Response {
  const bytes = new TextEncoder().encode(body)
  const size = opts.chunkBytes ?? 7
  let offset = 0
  const stream = new ReadableStream<Uint8Array>({
    pull(controller) {
      if (offset >= bytes.length) {
        controller.close()
        return
      }
      controller.enqueue(bytes.slice(offset, offset + size))
      offset += size
    },
  })
  return new Response(stream, {
    status: opts.status ?? 200,
    headers: { 'content-type': 'text/event-stream; charset=utf-8' },
  })
}

/** A stream whose body read fails mid-answer (connection reset). */
export function brokenSseResponse(prefix: string): Response {
  const bytes = new TextEncoder().encode(prefix)
  let sent = false
  const stream = new ReadableStream<Uint8Array>({
    pull(controller) {
      if (!sent) {
        sent = true
        controller.enqueue(bytes)
        return
      }
      controller.error(new TypeError('network error'))
    },
  })
  return new Response(stream, {
    status: 200,
    headers: { 'content-type': 'text/event-stream' },
  })
}

export function jsonErrorResponse(
  status: number,
  type: string,
  message: string,
  headers: Record<string, string> = {},
): Response {
  return new Response(JSON.stringify({ type: 'error', error: { type, message } }), {
    status,
    headers: { 'content-type': 'application/json', ...headers },
  })
}

/**
 * A complete answer the way a 5.x model streams it: a (display-omitted)
 * thinking block first, then the text split into several deltas.
 */
export function answerEvents(
  text: string,
  opts: {
    stopReason?: string
    stopDetails?: unknown
    usage?: { input: number; output: number; cacheWrite?: number; cacheRead?: number }
  } = {},
): SseEvent[] {
  const usage = opts.usage ?? { input: 120, output: 40, cacheWrite: 0, cacheRead: 0 }
  const pieces: string[] = []
  for (let i = 0; i < text.length; i += 9) pieces.push(text.slice(i, i + 9))
  return [
    {
      event: 'message_start',
      data: {
        type: 'message_start',
        message: {
          id: 'msg_fixture',
          type: 'message',
          role: 'assistant',
          model: 'claude-sonnet-5-5',
          content: [],
          stop_reason: null,
          usage: {
            input_tokens: usage.input,
            cache_creation_input_tokens: usage.cacheWrite ?? 0,
            cache_read_input_tokens: usage.cacheRead ?? 0,
            output_tokens: 1,
          },
        },
      },
    },
    { event: 'content_block_start', data: { type: 'content_block_start', index: 0, content_block: { type: 'thinking', thinking: '' } } },
    { event: 'content_block_delta', data: { type: 'content_block_delta', index: 0, delta: { type: 'signature_delta', signature: 'sig' } } },
    { event: 'content_block_stop', data: { type: 'content_block_stop', index: 0 } },
    { event: 'content_block_start', data: { type: 'content_block_start', index: 1, content_block: { type: 'text', text: '' } } },
    ...pieces.map((p) => ({
      event: 'content_block_delta',
      data: { type: 'content_block_delta', index: 1, delta: { type: 'text_delta', text: p } },
    })),
    { event: 'content_block_stop', data: { type: 'content_block_stop', index: 1 } },
    {
      event: 'message_delta',
      data: {
        type: 'message_delta',
        delta: {
          stop_reason: opts.stopReason ?? 'end_turn',
          stop_sequence: null,
          ...(opts.stopDetails !== undefined ? { stop_details: opts.stopDetails } : {}),
        },
        usage: { output_tokens: usage.output },
      },
    },
    { event: 'message_stop', data: { type: 'message_stop' } },
  ]
}

/**
 * A refusal as the API streams it (2026-11-06): `before-output` = no content
 * blocks at all; `after-thinking` = the 5.x thinking block streams first, then
 * the classifier declines before any answer text (a "mid-stream" refusal).
 */
export function refusalEvents(
  category: string | null,
  opts: { when?: 'before-output' | 'after-thinking'; input?: number; details?: Record<string, unknown> } = {},
): SseEvent[] {
  const events = answerEvents('', { usage: { input: opts.input ?? 300, output: 0 } })
  const start = events[0]!
  const thinking = events.slice(1, 4) // content_block_start / signature_delta / content_block_stop
  return [
    start,
    ...(opts.when === 'after-thinking' ? thinking : []),
    {
      event: 'message_delta',
      data: {
        type: 'message_delta',
        delta: {
          stop_reason: 'refusal',
          stop_sequence: null,
          stop_details: { type: 'refusal', category, explanation: null, ...opts.details },
        },
        usage: { output_tokens: 0 },
      },
    },
    { event: 'message_stop', data: { type: 'message_stop' } },
  ]
}

/** One classify answer for the given email indexes, as structured output. */
export function actionsJson(
  actions: Array<{ emailIndex: number; action?: string; targetFolderPath?: string; confidence?: number }>,
): string {
  return JSON.stringify({
    actions: actions.map((a) => ({
      emailIndex: a.emailIndex,
      reason: 'fixture',
      action: a.action ?? 'skip',
      ...(a.targetFolderPath ? { targetFolderPath: a.targetFolderPath } : {}),
      confidence: a.confidence ?? 0.9,
    })),
  })
}

/**
 * Recorded shape of a real classify stream (ids / text shortened): ping
 * keep-alives, a thinking block with its signature, CRLF line endings and a
 * text answer whose JSON is cut at arbitrary points across deltas.
 */
export const RECORDED_CLASSIFY_STREAM = [
  'event: message_start',
  'data: {"type":"message_start","message":{"id":"msg_01Rec","type":"message","role":"assistant","model":"claude-sonnet-5-5","content":[],"stop_reason":null,"stop_sequence":null,"usage":{"input_tokens":1834,"cache_creation_input_tokens":0,"cache_read_input_tokens":4210,"output_tokens":3}}}',
  '',
  'event: ping',
  'data: {"type": "ping"}',
  '',
  'event: content_block_start',
  'data: {"type":"content_block_start","index":0,"content_block":{"type":"thinking","thinking":"","signature":""}}',
  '',
  'event: content_block_delta',
  'data: {"type":"content_block_delta","index":0,"delta":{"type":"signature_delta","signature":"EqQBCkgIARABGAIiQL"}}',
  '',
  'event: content_block_stop',
  'data: {"type":"content_block_stop","index":0}',
  '',
  'event: content_block_start',
  'data: {"type":"content_block_start","index":1,"content_block":{"type":"text","text":""}}',
  '',
  'event: content_block_delta',
  'data: {"type":"content_block_delta","index":1,"delta":{"type":"text_delta","text":"{\\"actions\\":[{\\"emailIndex\\":0,\\"reason\\":\\"寄件網域 @client.example、主旨"}}',
  '',
  'event: ping',
  'data: {"type": "ping"}',
  '',
  'event: content_block_delta',
  'data: {"type":"content_block_delta","index":1,"delta":{"type":"text_delta","text":"提及 112訴204\\",\\"action\\":\\"move\\",\\"targetFolderPath\\":\\"內部資料/工時審閱\\",\\"confidence\\":0.92},"}}',
  '',
  'event: content_block_delta',
  'data: {"type":"content_block_delta","index":1,"delta":{"type":"text_delta","text":"{\\"emailIndex\\":1,\\"reason\\":\\"電子報\\",\\"action\\":\\"delete\\",\\"confidence\\":0.95}]}"}}',
  '',
  'event: content_block_stop',
  'data: {"type":"content_block_stop","index":1}',
  '',
  'event: message_delta',
  'data: {"type":"message_delta","delta":{"stop_reason":"end_turn","stop_sequence":null},"usage":{"output_tokens":187}}',
  '',
  'event: message_stop',
  'data: {"type":"message_stop"}',
  '',
  '',
].join('\r\n')
