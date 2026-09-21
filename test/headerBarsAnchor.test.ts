import assert from 'node:assert/strict'
import test from 'node:test'
import type { HeaderBarsSourceAnchorRect } from '../src/OfficeSDK.headerBars'
import { toHeaderBarsCallbackAnchorRect } from '../src/headerBarsAnchor'

/** iframe 视口 800x600 内的锚点矩形，用于换算断言。 */
function createSourceAnchorRect(
  overrides: Partial<HeaderBarsSourceAnchorRect> = {}
): HeaderBarsSourceAnchorRect {
  return {
    left: 100,
    top: 50,
    right: 140,
    bottom: 80,
    width: 40,
    height: 30,
    viewport: { width: 800, height: 600 },
    coordinateSpace: 'sdk-iframe-viewport',
    ...overrides
  }
}

void test('converts anchor rect with scale and offset', () => {
  const result = toHeaderBarsCallbackAnchorRect(
    createSourceAnchorRect(),
    { left: 200, top: 100, width: 400, height: 300 },
    { width: 1440, height: 900 }
  )

  // scaleX = 400/800 = 0.5, scaleY = 300/600 = 0.5
  assert.equal(result.left, 250)
  assert.equal(result.top, 125)
  assert.equal(result.right, 270)
  assert.equal(result.bottom, 140)
  assert.equal(result.width, 20)
  assert.equal(result.height, 15)
  assert.equal(result.coordinateSpace, 'callback-window-viewport')
  assert.deepEqual(result.viewport, { width: 1440, height: 900 })
})

void test('falls back to source viewport size when iframe rect is missing', () => {
  const result = toHeaderBarsCallbackAnchorRect(
    createSourceAnchorRect(),
    null,
    { width: 1024, height: 768 }
  )

  // 无 iframe 矩形时按 1:1 处理，仅套用来源矩形
  assert.equal(result.left, 100)
  assert.equal(result.top, 50)
  assert.equal(result.right, 140)
  assert.equal(result.bottom, 80)
  assert.equal(result.width, 40)
  assert.equal(result.height, 30)
})

void test('treats zero viewport size as one to avoid division by zero', () => {
  const result = toHeaderBarsCallbackAnchorRect(
    createSourceAnchorRect({ viewport: { width: 0, height: 0 } }),
    { left: 10, top: 20, width: 800, height: 600 },
    { width: 800, height: 600 }
  )

  assert.equal(result.left, 10 + 100 * 800)
  assert.equal(Number.isFinite(result.width), true)
  assert.equal(Number.isFinite(result.height), true)
})

void test('keeps integer geometry for an unscaled iframe', () => {
  const result = toHeaderBarsCallbackAnchorRect(
    createSourceAnchorRect(),
    { left: 0, top: 0, width: 800, height: 600 },
    { width: 1280, height: 800 }
  )

  assert.equal(result.left, 100)
  assert.equal(result.top, 50)
  assert.equal(result.width, 40)
  assert.equal(result.height, 30)
})
