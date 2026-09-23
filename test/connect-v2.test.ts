import assert from 'node:assert/strict'
import test from 'node:test'

async function loadConnectV2Api() {
  if (!(globalThis as { self?: unknown }).self) {
    ;(globalThis as { self?: unknown }).self = globalThis
  }
  return await import('../src')
}

void test('connectV2 exposes the supported file types', async () => {
  const {
    CONNECT_V2_FILE_TYPES,
    isConnectV2FileType,
    isConnectV2Mode,
    validateConnectV2Options
  } = await loadConnectV2Api()
  assert.deepEqual(CONNECT_V2_FILE_TYPES, [
    'file',
    'document',
    'documentPro',
    'writer',
    'spreadsheet',
    'presentation',
    'table',
    'form'
  ])
  assert.equal(isConnectV2FileType('writer'), true)
  assert.equal(isConnectV2FileType('unknown'), false)
  assert.equal(isConnectV2Mode('edit'), true)
  assert.equal(isConnectV2Mode('preview'), true)
  assert.equal(isConnectV2Mode('unknown'), false)
  assert.throws(
    () => validateConnectV2Options({ mode: 'edit' }),
    /invalid connectV2 "type"/
  )
  assert.throws(
    () => validateConnectV2Options({ type: 'writer', mode: 'unknown' }),
    /invalid connectV2 "mode"/
  )
})
