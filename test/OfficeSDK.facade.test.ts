import assert from 'node:assert/strict'
import test from 'node:test'
import { FileType } from 'weboffice-js-sdk-shared'
import { buildRootFacadeState } from '../src/OfficeSDK.facade'

type InvokeResponse =
  | unknown
  | ((args: unknown[]) => unknown | Promise<unknown>)

function createHost(fileType: FileType) {
  const calls: Array<{ method: string; args: unknown[] }> = []
  const responses = new Map<string, InvokeResponse>()
  const invokeEditorFacade = async <T>(
    method: string,
    args: unknown[] = []
  ): Promise<T> => {
    calls.push({ method, args })
    const response = responses.get(method)
    return (
      typeof response === 'function' ? await response(args) : response
    ) as T
  }
  const createValueObjectFacade = <T extends object>(
    prefix: string,
    locator: Record<string, unknown>,
    staticFields: Partial<T>
  ): T =>
    new Proxy(staticFields as T, {
      get(target, prop) {
        if (prop === 'then') return undefined
        if (Object.prototype.hasOwnProperty.call(target, prop)) {
          return target[prop as keyof T]
        }
        return async (...args: unknown[]) =>
          await invokeEditorFacade(`${prefix}.${String(prop)}`, [
            locator,
            ...args
          ])
      }
    })
  const createEditorFacadeModule = <T extends object>(
    prefix: string,
    staticFields: Partial<T> = {}
  ): T =>
    new Proxy(staticFields as T, {
      get(target, prop) {
        if (prop === 'then') return undefined
        if (Object.prototype.hasOwnProperty.call(target, prop)) {
          return target[prop as keyof T]
        }
        return async (...args: unknown[]) =>
          await invokeEditorFacade(`${prefix}.${String(prop)}`, args)
      }
    })

  return {
    calls,
    responses,
    host: {
      fileType,
      invokeEditorFacade,
      listenEditorEvent: () => () => undefined,
      createEditorFacadeModule,
      createValueObjectFacade,
      registerEditorFacadeListener: () => () => undefined,
      registerEditorFacadeCallback: () => 'callback-id',
      unregisterEditorFacadeCallback: () => undefined
    }
  }
}

void test('wraps SheetSelection locators and preserves optional arguments', async () => {
  const { host, responses, calls } = createHost(FileType.Spreadsheet)
  responses.set('sheet.worksheet.getSelections', [
    {
      kind: 'sheet.selection',
      sheetId: 'sheet-1',
      selectionId: 'selection-1'
    }
  ])
  responses.set('sheet.selection.getRange', {
    sheetId: 'sheet-1',
    row: 1,
    column: 2,
    rowCount: 3,
    columnCount: 4
  })
  responses.set('sheet.selection.setRange', undefined)

  const facade = buildRootFacadeState(host)
  const selections = await facade.activeSheet?.getSelections()
  assert.equal(selections?.length, 1)
  assert.equal('selectionId' in (selections?.[0] ?? {}), false)
  assert.equal('sheetId' in (selections?.[0] ?? {}), false)
  const range = await selections?.[0].getRange()
  assert.deepEqual(
    {
      sheetId: range?.sheetId,
      row: range?.row,
      column: range?.column,
      rowCount: range?.rowCount,
      columnCount: range?.columnCount
    },
    {
      sheetId: 'sheet-1',
      row: 1,
      column: 2,
      rowCount: 3,
      columnCount: 4
    }
  )
  await selections?.[0].setRange(null)

  assert.deepEqual(calls[1], {
    method: 'sheet.selection.getRange',
    args: [
      {
        kind: 'sheet.selection',
        sheetId: 'sheet-1',
        selectionId: 'selection-1'
      }
    ]
  })
  assert.equal(calls[2].method, 'sheet.selection.setRange')
  assert.equal(calls[2].args[1], null)
})

void test('wraps Presentation Table, Cell, and Range locators', async () => {
  const { host, responses, calls } = createHost(FileType.Presentation)
  responses.set('slides.getCurrentSlide', { slideId: 'slide-1' })
  responses.set('slides.slide.getTables', [
    {
      kind: 'presentation.table',
      slideId: 'slide-1',
      tableId: 'table-1',
      id: 'table-1',
      rowCount: 2,
      columnCount: 3
    }
  ])
  responses.set('slides.slide.table.getCell', {
    kind: 'presentation.table.cell',
    slideId: 'slide-1',
    tableId: 'table-1',
    row: 1,
    column: 2
  })
  responses.set('slides.slide.table.getRange', {
    kind: 'presentation.table.range',
    slideId: 'slide-1',
    tableId: 'table-1',
    range: { row: 0, column: 0, rowCount: 2, columnCount: 2 }
  })

  const facade = buildRootFacadeState(host)
  const slide = await facade.slides?.getCurrentSlide()
  const [table] = (await slide?.getTables()) ?? []
  assert.deepEqual(
    { id: table.id, rowCount: table.rowCount, columnCount: table.columnCount },
    { id: 'table-1', rowCount: 2, columnCount: 3 }
  )
  const cell = await table.getCell(1, 2)
  const range = await table.getRange({
    row: 0,
    column: 0,
    rowCount: 2,
    columnCount: 2
  })
  await cell?.clearStyle()
  await range?.setSpan()

  assert.ok(cell)
  assert.ok(range)
  assert.equal('slideId' in cell, false)
  assert.equal('tableId' in cell, false)
  assert.equal('slideId' in range, false)
  assert.equal('tableId' in range, false)

  assert.equal(calls.at(-2)?.method, 'slides.slide.table.cell.clearStyle')
  assert.equal(calls.at(-1)?.method, 'slides.slide.table.range.setSpan')
})

void test('preserves null and empty collection results', async () => {
  const { host, responses } = createHost(FileType.Spreadsheet)
  responses.set('sheet.worksheet.getSelections', null)
  const spreadsheet = buildRootFacadeState(host)
  assert.equal(await spreadsheet.activeSheet?.getSelections(), null)

  const presentationHost = createHost(FileType.Presentation)
  presentationHost.responses.set('slides.getCurrentSlide', {
    slideId: 'slide-1'
  })
  presentationHost.responses.set('slides.slide.getTables', [])
  const presentation = buildRootFacadeState(presentationHost.host)
  const slide = await presentation.slides?.getCurrentSlide()
  assert.deepEqual(await slide?.getTables(), [])
})

void test('keeps suite-specific presentation facade exposure', () => {
  const document = buildRootFacadeState(createHost(FileType.Document).host)
  const spreadsheet = buildRootFacadeState(
    createHost(FileType.Spreadsheet).host
  )
  const presentation = buildRootFacadeState(
    createHost(FileType.Presentation).host
  )

  assert.deepEqual(Object.keys(document.presentation ?? {}).sort(), [
    'quit',
    'start'
  ])
  assert.deepEqual(Object.keys(spreadsheet.presentation ?? {}).sort(), [
    'quit',
    'start'
  ])
  assert.deepEqual(Object.keys(presentation.presentation ?? {}).sort(), [
    'addChangeListener',
    'quit',
    'start',
    'startFromCurrent',
    'startRemoteLive',
    'startSpeakerView'
  ])
})

void test('passes existing invoke errors through without an envelope', async () => {
  const { host, responses } = createHost(FileType.Spreadsheet)
  const invokeError = Object.assign(new Error('iframe failed'), {
    name: 'InvokeError',
    method: 'InvokeEditorMethod',
    arguments: []
  })
  responses.set('sheet.worksheet.getSelections', async () => {
    throw invokeError
  })
  const facade = buildRootFacadeState(host)

  await assert.rejects(
    facade.activeSheet?.getSelections() as Promise<unknown>,
    {
      name: 'InvokeError',
      message: 'iframe failed'
    }
  )
})
