import assert from 'node:assert/strict'
import test from 'node:test'
import {
  isPointerEventOnIframe,
  resolveOutsidePointerEventName
} from '../src/outsideIframePointer'

void test('treats the iframe element itself as inside', () => {
  const iframe = {}
  assert.equal(isPointerEventOnIframe({ target: iframe }, iframe), true)
})

void test('treats composedPath containing the iframe as inside', () => {
  const iframe = {}
  const host = {}
  assert.equal(
    isPointerEventOnIframe(
      {
        target: host,
        composedPath: () => [host, iframe]
      },
      iframe
    ),
    true
  )
})

void test('treats other targets as outside', () => {
  const iframe = {}
  const outside = {}
  assert.equal(
    isPointerEventOnIframe(
      {
        target: outside,
        composedPath: () => [outside]
      },
      iframe
    ),
    false
  )
})

void test('treats missing composedPath as outside when target is not the iframe', () => {
  const iframe = {}
  assert.equal(isPointerEventOnIframe({ target: {} }, iframe), false)
})

void test('uses pointerdown when PointerEvent exists', () => {
  assert.equal(
    resolveOutsidePointerEventName({
      PointerEvent: function PointerEvent() {}
    }),
    'pointerdown'
  )
})

void test('falls back to mousedown when PointerEvent is missing', () => {
  assert.equal(resolveOutsidePointerEventName({}), 'mousedown')
})
