import assert from 'node:assert/strict'
import test from 'node:test'
import {
  LIGHT_DOC_P0_CONTRACT,
  validateJSAPIContractCatalog
} from '../src/jsapi'

void test('light-document API contract catalog is internally consistent', () => {
  assert.equal(LIGHT_DOC_P0_CONTRACT.apis.length, 24)
  assert.deepEqual(validateJSAPIContractCatalog(LIGHT_DOC_P0_CONTRACT), [])
})

void test('contract catalog covers all communication and validation modes', () => {
  const transports = new Set(
    LIGHT_DOC_P0_CONTRACT.apis.map((api) => api.transport)
  )
  assert.deepEqual([...transports].sort(), [
    'delegation',
    'request',
    'subscription'
  ])
  assert.ok(
    LIGHT_DOC_P0_CONTRACT.apis.some((api) => api.adapter === 'delta-snapshot')
  )
  assert.ok(
    LIGHT_DOC_P0_CONTRACT.apis.some((api) => api.adapter === 'range-value')
  )
  assert.ok(
    LIGHT_DOC_P0_CONTRACT.apis.some((api) =>
      api.sideEffects.includes('selection-style')
    )
  )
})
