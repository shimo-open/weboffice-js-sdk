import assert from 'node:assert/strict'
import test from 'node:test'
import { createPreloadInitPayload } from '../src/preload'

const baseOptions = {
  token: 'token-1',
  signature: 'signature-1',
  fileGuid: 'file-guid',
  mode: 'edit' as const,
  type: 'writer' as const
}

void test('preload init payload marks only explicit V2 connections', () => {
  assert.deepEqual(
    createPreloadInitPayload({
      ...baseOptions,
      contentVersion: 'v2'
    }),
    {
      ...baseOptions,
      contentVersion: 'v2'
    }
  )
})

void test('preload init payload omits incomplete V2 markers', () => {
  for (const contentVersion of [undefined, 'v1', 'V2', true] as unknown[]) {
    assert.deepEqual(
      createPreloadInitPayload({
        ...baseOptions,
        contentVersion: contentVersion as never
      }),
      baseOptions
    )
  }
})
