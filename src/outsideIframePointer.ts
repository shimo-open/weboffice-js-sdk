export const OUTSIDE_POINTER_DOWN_METHOD = 'sdk.outsidePointerDown'

/**
 * 判断父页按下是否落在指定 iframe 元素上（含 Shadow DOM composedPath）。
 * 输入：指针事件与 iframe 元素。
 * 输出：true 表示点在这张 iframe 上，宿主不应通知结束编辑。
 */
export function isPointerEventOnIframe(
  event: { target?: unknown; composedPath?: () => unknown[] },
  iframe: unknown
): boolean {
  if (event.target === iframe) {
    return true
  }
  if (typeof event.composedPath !== 'function') {
    return false
  }
  try {
    return event.composedPath().includes(iframe)
  } catch {
    return false
  }
}

/**
 * 选择外层「点在 iframe 外」监听的事件名。
 * 输入：可选的全局对象，默认当前 globalThis。
 * 输出：支持 PointerEvent 时用 pointerdown，否则回退 mousedown。
 */
export function resolveOutsidePointerEventName(
  globalObject: {
    PointerEvent?: unknown
  } = globalThis
): 'pointerdown' | 'mousedown' {
  return typeof globalObject.PointerEvent === 'function'
    ? 'pointerdown'
    : 'mousedown'
}
