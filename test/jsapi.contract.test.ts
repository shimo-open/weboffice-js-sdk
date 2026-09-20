import assert from 'node:assert/strict'
import test from 'node:test'
import {
  LIGHT_DOC_P0_CONTRACT,
  PUBLIC_METHODS_CONTRACT,
  validateJSAPIContractCatalog
} from '../src/jsapi'
import { OfficeSDKMethods } from '../src/OfficeSDK.methods'

void test('light-document API contract catalog is internally consistent', () => {
  assert.equal(LIGHT_DOC_P0_CONTRACT.apis.length, 23)
  assert.deepEqual(validateJSAPIContractCatalog(LIGHT_DOC_P0_CONTRACT), [])
})

void test('public capability methods stay aligned with the discoverable SDK constants', () => {
  const methodPaths = new Set(
    PUBLIC_METHODS_CONTRACT.methods.map((method) => method.publicPath)
  )
  const constantPaths = new Set<string>()
  const visit = (value: unknown) => {
    if (typeof value === 'string') {
      constantPaths.add(value)
      return
    }
    if (value && typeof value === 'object') {
      Object.values(value).forEach(visit)
    }
  }
  visit(OfficeSDKMethods)
  assert.deepEqual(constantPaths, methodPaths)
  assert.equal(PUBLIC_METHODS_CONTRACT.methods.length, 23)
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
