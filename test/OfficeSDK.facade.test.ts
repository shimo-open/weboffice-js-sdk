import assert from 'node:assert/strict'
import test from 'node:test'
import { FileType } from 'weboffice-js-sdk-shared'
import { buildRootFacadeState } from '../src/OfficeSDK.facade'
import {
  HEADER_BARS_METHOD,
  initHeaderBarsFacade,
  type HeaderBarsCapabilities,
  type HeaderBarsCommandRef,
  type HeaderBarsCommandState
} from '../src/OfficeSDK.headerBars'

type InvokeResponse =
  | unknown
  | ((args: unknown[]) => unknown | Promise<unknown>)

function createHost(fileType: FileType) {
  const calls: Array<{ method: string; args: unknown[] }> = []
  const responses = new Map<string, InvokeResponse>()
  const callbacks = new Map<
    string,
    (...args: unknown[]) => unknown | Promise<unknown>
  >()
  const errors: Array<{ message: string; error: unknown }> = []
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
        if (
          prop === 'then' ||
          prop === 'catch' ||
          prop === 'finally' ||
          prop === 'toJSON'
        ) {
          return undefined
        }
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
        if (
          prop === 'then' ||
          prop === 'catch' ||
          prop === 'finally' ||
          prop === 'toJSON'
        ) {
          return undefined
        }
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
    callbacks,
    errors,
    host: {
      fileType,
      invokeEditorFacade,
      listenEditorEvent: () => () => undefined,
      createEditorFacadeModule,
      createValueObjectFacade,
      registerEditorFacadeListener: () => () => undefined,
      registerEditorFacadeCallback: (
        callback: (...args: unknown[]) => unknown | Promise<unknown>
      ) => {
        callbacks.set('callback-id', callback)
        return 'callback-id'
      },
      unregisterEditorFacadeCallback: (callbackId: string) => {
        callbacks.delete(callbackId)
      },
      reportEditorFacadeError: (message: string, error: unknown) => {
        errors.push({ message, error })
      }
    }
  }
}

async function flushPromises() {
  await new Promise<void>((resolve) => setImmediate(resolve))
}

void test('mounts the typed ActiveOutline facade and document roots only for documents', async () => {
  const document = buildRootFacadeState(createHost(FileType.Document).host)
  const spreadsheet = buildRootFacadeState(
    createHost(FileType.Spreadsheet).host
  )
  const presentation = buildRootFacadeState(
    createHost(FileType.Presentation).host
  )

  assert.ok(document.ActiveOutline)
  assert.ok(document.ActiveOutline?.Reference)
  assert.ok(document.ActiveOutline?.Service)
  assert.ok(document.ActiveOutline?.Sub)
  assert.ok(document.ActiveOutline?.Env)
  assert.equal('ActiveDocument' in document, false)
  assert.equal(spreadsheet.ActiveOutline, undefined)
  assert.equal(presentation.ActiveOutline, undefined)
})

void test('routes all typed root methods through productJSAPI paths', async () => {
  const { host, responses, calls } = createHost(FileType.Document)
  const prefix = 'productJSAPI.'
  responses.set(`${prefix}Editor.Document.GetContent`, {
    length: 3,
    serialized: '{"ops":[{"insert":"Hi\\n"}]}'
  })
  responses.set(`${prefix}Editor.Document.GetTitleContent`, 'Title')
  responses.set(`${prefix}Editor.GetEditMode`, 'edit')
  responses.set(`${prefix}Editor.Document.Font.SetTextColor`, true)
  responses.set(`${prefix}Editor.Document.Font.SetHighLightColor`, true)
  responses.set(`${prefix}Editor.Document.Font.SetBold`, true)
  responses.set(`${prefix}Editor.Document.Font.SetItalic`, true)
  responses.set(`${prefix}Editor.Document.Font.SetUnderline`, true)
  responses.set(`${prefix}Editor.Document.Font.SetStrike`, true)
  responses.set(`${prefix}Reference.CanIUse`, true)
  responses.set(`${prefix}Service.User.GetUserInfo`, { id: 'user-1' })
  responses.set(`${prefix}Service.Permission.GetDocumentPermission`, {
    read: true,
    write: true,
    comment: false
  })
  responses.set(`${prefix}Service.Collaboration.GetSaveStatus`, 'saved')
  responses.set(`${prefix}Editor.Document.Markdown.GetMarkdown`, '# Title')
  responses.set(`${prefix}Editor.Document.Markdown.AppendMarkdown`, {
    start: 1,
    end: 2
  })
  responses.set(`${prefix}Editor.Document.Markdown.InsertMarkdown`, {
    start: 3,
    end: 4
  })
  responses.set(`${prefix}Editor.Document.Markdown.ValidateMarkdown`, true)
  responses.set(`${prefix}Editor.Document.Content.ReplaceSelection`, true)
  responses.set(`${prefix}Editor.Document.Content.ReplaceAllContent`, true)
  responses.set(`${prefix}Env.Language.GetLanguage`, 'zh-CN')
  responses.set(`${prefix}Env.DocsMode.GetDocsMode`, 'normal')

  const root = buildRootFacadeState(host)
  assert.ok(root.ActiveOutline)
  assert.ok(root.ActiveOutline.Reference)
  assert.ok(root.ActiveOutline.Service)
  assert.ok(root.ActiveOutline.Sub)
  assert.ok(root.ActiveOutline.Env)
  const editor = root.ActiveOutline.Editor
  const document = editor.Document
  const snapshot = await document.GetContent()
  assert.deepEqual(
    {
      length: snapshot.length,
      serialized: snapshot.serialized,
      stringified: snapshot.stringify()
    },
    {
      length: 3,
      serialized: '{"ops":[{"insert":"Hi\\n"}]}',
      stringified: '{"ops":[{"insert":"Hi\\n"}]}'
    }
  )
  assert.equal('compose' in snapshot, false)
  assert.equal('transform' in snapshot, false)
  assert.equal(await document.GetTitleContent(), 'Title')
  await document.SetTitleContent('Next title')
  assert.equal(await editor.GetEditMode(), 'edit')
  assert.equal(await document.Font.SetTextColor('#000000'), true)
  assert.equal(await document.Font.SetHighLightColor('#ffff00'), true)
  assert.equal(await document.Font.SetBold(), true)
  assert.equal(await document.Font.SetItalic(false), true)
  assert.equal(await document.Font.SetUnderline(true), true)
  assert.equal(await document.Font.SetStrike(false), true)
  assert.equal(await root.ActiveOutline.Reference.CanIUse(['a', 'b']), true)
  assert.deepEqual(await root.ActiveOutline.Service.User.GetUserInfo(), {
    id: 'user-1'
  })
  assert.deepEqual(
    await root.ActiveOutline.Service.Permission.GetDocumentPermission(),
    { read: true, write: true, comment: false }
  )
  assert.equal(
    await root.ActiveOutline.Service.Collaboration.GetSaveStatus(),
    'saved'
  )
  assert.equal(await document.Markdown.GetMarkdown(), '# Title')
  assert.deepEqual(await document.Markdown.AppendMarkdown('tail'), {
    start: 1,
    end: 2
  })
  assert.deepEqual(await document.Markdown.InsertMarkdown('middle'), {
    start: 3,
    end: 4
  })
  assert.equal(await document.Markdown.ValidateMarkdown('# ok'), true)
  assert.equal(await document.Content.ReplaceSelection('selection'), true)
  assert.equal(await document.Content.ReplaceAllContent('all'), true)
  assert.equal(await root.ActiveOutline.Env.Language.GetLanguage(), 'zh-CN')
  assert.equal(await root.ActiveOutline.Env.DocsMode.GetDocsMode(), 'normal')
  await root.ActiveOutline.Service.Export.DownloadDocument('pdf')

  assert.deepEqual(
    calls.map(({ method, args }) => ({ method, args })),
    [
      { method: `${prefix}Editor.Document.GetContent`, args: [] },
      { method: `${prefix}Editor.Document.GetTitleContent`, args: [] },
      {
        method: `${prefix}Editor.Document.SetTitleContent`,
        args: ['Next title']
      },
      { method: `${prefix}Editor.GetEditMode`, args: [] },
      {
        method: `${prefix}Editor.Document.Font.SetTextColor`,
        args: ['#000000']
      },
      {
        method: `${prefix}Editor.Document.Font.SetHighLightColor`,
        args: ['#ffff00']
      },
      { method: `${prefix}Editor.Document.Font.SetBold`, args: [] },
      {
        method: `${prefix}Editor.Document.Font.SetItalic`,
        args: [false]
      },
      {
        method: `${prefix}Editor.Document.Font.SetUnderline`,
        args: [true]
      },
      {
        method: `${prefix}Editor.Document.Font.SetStrike`,
        args: [false]
      },
      { method: `${prefix}Reference.CanIUse`, args: [['a', 'b']] },
      { method: `${prefix}Service.User.GetUserInfo`, args: [] },
      {
        method: `${prefix}Service.Permission.GetDocumentPermission`,
        args: []
      },
      {
        method: `${prefix}Service.Collaboration.GetSaveStatus`,
        args: []
      },
      {
        method: `${prefix}Editor.Document.Markdown.GetMarkdown`,
        args: []
      },
      {
        method: `${prefix}Editor.Document.Markdown.AppendMarkdown`,
        args: ['tail']
      },
      {
        method: `${prefix}Editor.Document.Markdown.InsertMarkdown`,
        args: ['middle']
      },
      {
        method: `${prefix}Editor.Document.Markdown.ValidateMarkdown`,
        args: ['# ok']
      },
      {
        method: `${prefix}Editor.Document.Content.ReplaceSelection`,
        args: ['selection']
      },
      {
        method: `${prefix}Editor.Document.Content.ReplaceAllContent`,
        args: ['all']
      },
      { method: `${prefix}Env.Language.GetLanguage`, args: [] },
      { method: `${prefix}Env.DocsMode.GetDocsMode`, args: [] },
      {
        method: `${prefix}Service.Export.DownloadDocument`,
        args: ['pdf']
      }
    ]
  )
})

void test('rebuilds document change snapshots and disposes after async registration', async () => {
  const { host, responses, calls, callbacks } = createHost(FileType.Document)
  let resolveRegistration: ((subscriptionId: string) => void) | undefined
  responses.set(
    'productJSAPI.Sub.OnDocumentChange',
    async () =>
      await new Promise<string>((resolve) => {
        resolveRegistration = resolve
      })
  )
  const received: Array<{
    length: number
    serialized: string
    stringified: string
  }> = []
  const root = buildRootFacadeState(host)
  const activeOutline = root.ActiveOutline
  assert.ok(activeOutline)
  const dispose = activeOutline.Sub.OnDocumentChange((snapshot) => {
    received.push({
      length: snapshot.length,
      serialized: snapshot.serialized,
      stringified: snapshot.stringify()
    })
  })

  await callbacks.get('callback-id')?.({ length: 2, serialized: 'delta-1' })
  assert.deepEqual(received, [
    { length: 2, serialized: 'delta-1', stringified: 'delta-1' }
  ])
  dispose()
  dispose()
  await callbacks.get('callback-id')?.({ length: 3, serialized: 'ignored' })
  assert.equal(received.length, 1)
  assert.equal(
    calls.some(({ method }) => method === 'productJSAPI.Sub.OffDocumentChange'),
    false
  )

  resolveRegistration?.('subscription-1')
  await flushPromises()
  assert.deepEqual(calls.at(-1), {
    method: 'productJSAPI.Sub.OffDocumentChange',
    args: ['subscription-1']
  })
  assert.equal(callbacks.has('callback-id'), false)
})

void test('reports document change registration and disposal failures', async () => {
  const registration = createHost(FileType.Document)
  const registrationError = new Error('registration failed')
  registration.responses.set('productJSAPI.Sub.OnDocumentChange', async () => {
    throw registrationError
  })
  const disposeRegistration = buildRootFacadeState(
    registration.host
  ).ActiveOutline?.Sub.OnDocumentChange(() => undefined)
  await flushPromises()
  assert.equal(registration.callbacks.size, 0)
  assert.deepEqual(registration.errors, [
    {
      message: 'register document change listener failed',
      error: registrationError
    }
  ])
  disposeRegistration?.()

  const disposal = createHost(FileType.Document)
  const disposalError = new Error('disposal failed')
  disposal.responses.set('productJSAPI.Sub.OnDocumentChange', 'subscription-1')
  disposal.responses.set('productJSAPI.Sub.OffDocumentChange', async () => {
    throw disposalError
  })
  const dispose = buildRootFacadeState(
    disposal.host
  ).ActiveOutline?.Sub.OnDocumentChange(() => undefined)
  dispose?.()
  await flushPromises()
  assert.equal(disposal.callbacks.size, 0)
  assert.deepEqual(disposal.errors, [
    {
      message: 'dispose document change listener failed',
      error: disposalError
    }
  ])
})

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
  assert.equal(selections?.[0].id, 'selection-1')
  assert.equal('selectionId' in (selections?.[0] ?? {}), false)
  assert.equal('sheetId' in (selections?.[0] ?? {}), false)
  assert.equal(JSON.stringify(selections?.[0]), '{"id":"selection-1"}')
  assert.equal(calls.length, 1)
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

void test('routes SheetRange getBounding through its range locator', async () => {
  const { host, responses, calls } = createHost(FileType.Spreadsheet)
  responses.set('sheet.worksheet.getRange', {
    sheetId: 'sheet-1',
    row: 1,
    column: 2,
    rowCount: 3,
    columnCount: 4
  })
  responses.set('sheet.range.getBounding', {
    left: 10,
    top: 20,
    width: 300,
    height: 120
  })

  const facade = buildRootFacadeState(host)
  const range = await facade.activeSheet?.getRange({
    type: 'cells',
    row: 1,
    column: 2,
    rowCount: 3,
    columnCount: 4
  })
  const bounding = await range?.getBounding()

  assert.deepEqual(bounding, {
    left: 10,
    top: 20,
    width: 300,
    height: 120
  })
  assert.equal(calls.at(-1)?.method, 'sheet.range.getBounding')
  assert.deepEqual(calls.at(-1)?.args, [
    {
      sheetId: 'sheet-1',
      row: 1,
      column: 2,
      rowCount: 3,
      columnCount: 4
    }
  ])
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

void test('HeaderBars extension negotiates capabilities and sends atomic batches', async () => {
  const calls: Array<{ method: string; payload?: Record<string, unknown> }> = []
  const capabilities: HeaderBarsCapabilities = {
    protocolVersion: 2,
    features: {
      treeCommands: true,
      batchCommands: true,
      sectionVisibility: true,
      commandOptions: true,
      commandOpenState: true,
      commandEvents: true,
      anchorEvents: true
    }
  }
  const commands = new Map<string, HeaderBarsCommandState>()
  const refs = new Map<string, HeaderBarsCommandRef>()
  const clickCallbacks = new Map<
    string,
    (() => void | Promise<void>) | undefined
  >()
  const eventCallbacks = new Map<
    string,
    ((...args: any[]) => void | Promise<void>) | undefined
  >()
  let cachedCapabilities: HeaderBarsCapabilities = {
    protocolVersion: 1,
    features: {
      treeCommands: false,
      batchCommands: false,
      sectionVisibility: false,
      commandOptions: false,
      commandOpenState: false,
      commandEvents: false,
      anchorEvents: false
    }
  }
  const facade = initHeaderBarsFacade({
    getVisibleState: () => true,
    setVisibleState: () => undefined,
    getCommandsMap: () => commands,
    getOverridesMap: () => clickCallbacks,
    getEventOverridesMap: () => eventCallbacks,
    getCapabilitiesState: () => cachedCapabilities,
    setCapabilitiesState: (next) => {
      cachedCapabilities = next
    },
    getIframeElement: () => null,
    getRefsMap: () => refs,
    getTitleHandler: () => undefined,
    setTitleHandler: () => undefined,
    isTitleSubscribed: () => false,
    setTitleSubscribed: () => undefined,
    invokeHeaderBars: async <T>(
      method: string,
      payload?: Record<string, unknown>
    ): Promise<T> => {
      calls.push({ method, payload })
      if (method === HEADER_BARS_METHOD.getCapabilities) {
        return capabilities as T
      }
      if (method === HEADER_BARS_METHOD.addCommands) {
        return JSON.parse('{"success":true}')
      }
      return undefined as T
    },
    emitHeaderBarsError: () => undefined,
    onInternalTitleChange: () => undefined,
    subscribeEditorTitleChange: async () => await Promise.resolve()
  })

  const onClick = async () => await Promise.resolve()
  const result = await facade.addCommands(
    [
      {
        id: 'custom-menu',
        section: 'more',
        onClick,
        subItems: [{ id: 'custom-child', label: 'Child' }]
      }
    ],
    'download',
    'before'
  )

  assert.deepEqual(result, { success: true })
  assert.equal(calls[0]?.method, HEADER_BARS_METHOD.getCapabilities)
  assert.equal(calls[1]?.method, HEADER_BARS_METHOD.addCommands)
  assert.equal(
    typeof (
      (calls[1]?.payload?.commands as Array<Record<string, unknown>>)?.[0] ?? {}
    ).onClick,
    'undefined'
  )
  assert.equal(eventCallbacks.get('custom-menu:click'), onClick)
})

/** V2 事件回调测试用的宿主桩，记录 invoke 与回调表。 */
function createHeaderBarsEventHost() {
  const calls: Array<{ method: string; payload?: Record<string, unknown> }> = []
  const commands = new Map<string, HeaderBarsCommandState>()
  const refs = new Map<string, HeaderBarsCommandRef>()
  const clickCallbacks = new Map<
    string,
    (() => void | Promise<void>) | undefined
  >()
  const eventCallbacks = new Map<
    string,
    ((...args: any[]) => void | Promise<void>) | undefined
  >()
  const capabilities: HeaderBarsCapabilities = {
    protocolVersion: 2,
    features: {
      treeCommands: true,
      batchCommands: true,
      sectionVisibility: true,
      commandOptions: true,
      commandOpenState: true,
      commandEvents: true,
      anchorEvents: true
    }
  }
  const facade = initHeaderBarsFacade({
    getVisibleState: () => true,
    setVisibleState: () => undefined,
    getCommandsMap: () => commands,
    getOverridesMap: () => clickCallbacks,
    getEventOverridesMap: () => eventCallbacks,
    getCapabilitiesState: () => capabilities,
    setCapabilitiesState: () => undefined,
    getIframeElement: () => null,
    getRefsMap: () => refs,
    getTitleHandler: () => undefined,
    setTitleHandler: () => undefined,
    isTitleSubscribed: () => false,
    setTitleSubscribed: () => undefined,
    invokeHeaderBars: async <T>(
      method: string,
      payload?: Record<string, unknown>
    ): Promise<T> => {
      calls.push({ method, payload })
      return undefined as T
    },
    emitHeaderBarsError: () => undefined,
    onInternalTitleChange: () => undefined,
    subscribeEditorTitleChange: async () => await Promise.resolve()
  })
  return { facade, calls, eventCallbacks, clickCallbacks }
}

void test('HeaderBars V2 registers open/close callbacks and toggles the iframe flag', async () => {
  const { facade, calls, eventCallbacks } = createHeaderBarsEventHost()
  const onOpen = async () => await Promise.resolve()
  const onClose = async () => await Promise.resolve()

  const ref = facade.getCommand('collaborators')
  ref.onCommandOpen = onOpen
  ref.onCommandClose = onClose

  assert.equal(eventCallbacks.get('collaborators:open'), onOpen)
  assert.equal(eventCallbacks.get('collaborators:close'), onClose)

  const openCalls = calls.filter(
    (call) =>
      call.method === HEADER_BARS_METHOD.setCommandEventCallbackEnabled &&
      call.payload?.event === 'open'
  )
  const closeCalls = calls.filter(
    (call) =>
      call.method === HEADER_BARS_METHOD.setCommandEventCallbackEnabled &&
      call.payload?.event === 'close'
  )
  assert.equal(openCalls.length, 1)
  assert.deepEqual(openCalls[0]?.payload, {
    id: 'collaborators',
    event: 'open',
    enabled: true
  })
  assert.equal(closeCalls.length, 1)
  assert.deepEqual(closeCalls[0]?.payload, {
    id: 'collaborators',
    event: 'close',
    enabled: true
  })

  // 解除注册后必须显式关闭 iframe 侧转发，否则默认成员卡片会被误抑制
  ref.onCommandOpen = undefined
  const disabledCalls = calls.filter(
    (call) =>
      call.method === HEADER_BARS_METHOD.setCommandEventCallbackEnabled &&
      call.payload?.event === 'open' &&
      call.payload?.enabled === false
  )
  assert.equal(disabledCalls.length, 1)
  assert.equal(eventCallbacks.get('collaborators:open'), undefined)
})

void test('HeaderBars V2 keeps click and open/close callbacks independent', async () => {
  const { facade, calls, eventCallbacks, clickCallbacks } =
    createHeaderBarsEventHost()
  const onClick = async () => await Promise.resolve()
  const onOpen = async () => await Promise.resolve()

  const ref = facade.getCommand('collaborators')
  ref.onCommandClick = onClick
  ref.onCommandOpen = onOpen

  // click 走 overrides map，open 走 event overrides map，两者不得互相污染
  assert.equal(clickCallbacks.get('collaborators'), onClick)
  assert.equal(eventCallbacks.get('collaborators:open'), onOpen)
  assert.equal(eventCallbacks.get('collaborators'), undefined)
  assert.equal(clickCallbacks.get('collaborators:open'), undefined)

  // click 走旧通道（setCommandCallbackEnabled），open 走新通道（setCommandEventCallbackEnabled）；
  // click 不得被写成事件回调，否则会额外开启 handleCommandEvent 转发
  const clickEventFlagCalls = calls.filter(
    (call) =>
      call.method === HEADER_BARS_METHOD.setCommandEventCallbackEnabled &&
      call.payload?.event === 'click'
  )
  assert.equal(clickEventFlagCalls.length, 0)
  assert.deepEqual(
    calls.map((call) => call.method),
    [
      HEADER_BARS_METHOD.setCommandCallbackEnabled,
      HEADER_BARS_METHOD.setCommandEventCallbackEnabled
    ]
  )
})
