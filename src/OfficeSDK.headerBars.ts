import { InvokeMethod } from 'weboffice-js-sdk-shared'

export const HEADER_BARS_METHOD = {
  getCapabilities: 'headerBars.getCapabilities',
  getVisible: 'headerBars.getVisible',
  setVisible: 'headerBars.setVisible',
  addCommand: 'headerBars.addCommand',
  getCommand: 'headerBars.getCommand',
  setCommandVisible: 'headerBars.setCommandVisible',
  setCommandDisabled: 'headerBars.setCommandDisabled',
  setCommandActive: 'headerBars.setCommandActive',
  setCommandSrc: 'headerBars.setCommandSrc',
  setCommandLabel: 'headerBars.setCommandLabel',
  setCommandEditable: 'headerBars.setCommandEditable',
  setCommandCallbackEnabled: 'headerBars.setCommandCallbackEnabled',
  listViewCommands: 'headerBars.listViewCommands',
  handleCommandClick: 'headerBars.handleCommandClick',
  addCommands: 'headerBars.addCommands',
  setSectionVisible: 'headerBars.setSectionVisible',
  setCommandOptions: 'headerBars.setCommandOptions',
  setCommandOpen: 'headerBars.setCommandOpen',
  setCommandEventCallbackEnabled: 'headerBars.setCommandEventCallbackEnabled',
  handleCommandEvent: 'headerBars.handleCommandEvent'
} as const

export const HEADER_BARS_CHANGED_EVENT = 'headerBars:changed'

export interface HeaderBarsCommandDefinition {
  id: string
  section?: string
  order?: number
  label?: string
  visible?: boolean
  disabled?: boolean
  active?: boolean
  editable?: boolean
  type?: 'action' | 'structural'
  renderType?: string
  src?: string
  danger?: boolean
  disabledTip?: string
  itemType?: 'command' | 'divider'
  trigger?: 'hover' | 'click'
  open?: boolean
  style?: HeaderBarsCommandStyle
  toast?: Record<string, unknown>
  subItems?: HeaderBarsCommandDefinition[]
  onClick?: () => void | Promise<void>
  onOpen?: HeaderBarsCommandOpenHandler
  onClose?: HeaderBarsCommandCloseHandler
}

export interface HeaderBarsCommandState extends HeaderBarsCommandDefinition {
  type: 'action' | 'structural'
  open?: boolean
}

export interface HeaderBarsCommandRef {
  readonly id: string
  visible: boolean
  disabled: boolean
  active: boolean
  src?: string
  label?: string
  editable?: boolean
  open: boolean
  style?: HeaderBarsCommandStyle
  subItems?: HeaderBarsCommandDefinition[]
  onCommandClick?: () => void | Promise<void>
  onCommandOpen?: HeaderBarsCommandOpenHandler
  onCommandClose?: HeaderBarsCommandCloseHandler
  getState: () => HeaderBarsCommandState | undefined
}

export interface HeaderBarsCommandStyle {
  color?: string
  fontSize?: number
  fontWeight?: number
  backgroundColor?: string
  borderColor?: string
  borderWidth?: number
  borderRadius?: number
  width?: number
  height?: number
}

export interface HeaderBarsAnchorRect {
  left: number
  top: number
  right: number
  bottom: number
  width: number
  height: number
  viewport: { width: number; height: number }
  coordinateSpace: 'callback-window-viewport'
}

export interface HeaderBarsSourceAnchorRect {
  left: number
  top: number
  right: number
  bottom: number
  width: number
  height: number
  viewport: { width: number; height: number }
  coordinateSpace: 'sdk-iframe-viewport'
}

export interface HeaderBarsCommandEventPayload {
  requestId: string
  sequence: number
  commandId: string
  event: 'click' | 'open' | 'close'
  anchorRect: HeaderBarsSourceAnchorRect
  context?: Record<string, unknown>
}

export type HeaderBarsCommandEventResult =
  | 'handled'
  | 'unhandled'
  | 'failed'
  | 'timedOut'

export interface HeaderBarsCommandEvent {
  commandId: string
  event: 'click' | 'open' | 'close'
  anchorRect: HeaderBarsAnchorRect
  context?: Record<string, unknown>
}

export type HeaderBarsCommandClickHandler = (
  event?: HeaderBarsCommandEvent
) => void | Promise<void>
export type HeaderBarsCommandOpenHandler = (
  commandId: string,
  anchorRect: HeaderBarsAnchorRect,
  context: Record<string, unknown>
) => void | Promise<void>
export type HeaderBarsCommandCloseHandler = HeaderBarsCommandOpenHandler
type HeaderBarsEventHandler = (...args: any[]) => void | Promise<void>

export interface HeaderBarsCapabilities {
  protocolVersion: 1 | 2
  features: {
    treeCommands: boolean
    batchCommands: boolean
    sectionVisibility: boolean
    commandOptions: boolean
    commandOpenState: boolean
    commandEvents: boolean
    anchorEvents: boolean
  }
}

export interface HeaderBarsMutationResult {
  success: boolean
  code?: string
  message?: string
}

export type HeaderBarsTitleChangeHandler = (
  title: string
) => void | Promise<void>

export interface HeaderBarsFacade {
  visible: boolean
  onTitleChange?: HeaderBarsTitleChangeHandler
  getVisible: () => Promise<boolean>
  setVisible: (visible: boolean) => Promise<void>
  addCommand: (
    command: HeaderBarsCommandDefinition,
    posCommand: string,
    pos?: 'before' | 'after'
  ) => Promise<boolean>
  getCommand: (id: string) => HeaderBarsCommandRef
  listViewCommands: () => Promise<HeaderBarsCommandState[]>
  getCapabilities: () => Promise<HeaderBarsCapabilities>
  addCommands: (
    commands: HeaderBarsCommandDefinition[],
    posCommand: string,
    pos?: 'before' | 'after'
  ) => Promise<HeaderBarsMutationResult>
  setSectionVisible: (
    section: 'left' | 'right',
    visible: boolean
  ) => Promise<HeaderBarsMutationResult>
}

export interface HeaderBarsChangedPayload {
  reason?: string
  commandId?: string
  version?: number
  snapshot?: {
    visible: boolean
    commands: HeaderBarsCommandState[]
  }
}

interface HeaderBarsHost {
  getVisibleState(): boolean
  setVisibleState(visible: boolean): void
  getCommandsMap(): Map<string, HeaderBarsCommandState>
  getOverridesMap(): Map<string, (() => void | Promise<void>) | undefined>
  getEventOverridesMap(): Map<string, HeaderBarsEventHandler | undefined>
  getCapabilitiesState(): HeaderBarsCapabilities
  setCapabilitiesState(capabilities: HeaderBarsCapabilities): void
  getIframeElement(): HTMLElement | null
  getRefsMap(): Map<string, HeaderBarsCommandRef>
  getTitleHandler(): HeaderBarsTitleChangeHandler | undefined
  setTitleHandler(handler: HeaderBarsTitleChangeHandler | undefined): void
  isTitleSubscribed(): boolean
  setTitleSubscribed(subscribed: boolean): void
  invokeHeaderBars<T>(
    method: string,
    payload?: Record<string, unknown>
  ): Promise<T>
  emitHeaderBarsError(message: string, err: unknown): void
  onInternalTitleChange(listener: (title: unknown) => void): void
  subscribeEditorTitleChange(): Promise<void>
}

function serializeCommandDefinitions(commands: HeaderBarsCommandDefinition[]) {
  const callbacks = new Map<string, HeaderBarsEventHandler | undefined>()

  const serialize = (
    command: HeaderBarsCommandDefinition
  ): HeaderBarsCommandDefinition => {
    if (typeof command.onClick === 'function') {
      callbacks.set(`${command.id}:click`, command.onClick)
    }
    if (typeof command.onOpen === 'function') {
      callbacks.set(`${command.id}:open`, command.onOpen)
    }
    if (typeof command.onClose === 'function') {
      callbacks.set(`${command.id}:close`, command.onClose)
    }
    const { onClick, onOpen, onClose, subItems, ...payload } = command
    void onClick
    void onOpen
    void onClose
    return {
      ...payload,
      subItems: subItems?.map(serialize)
    }
  }

  return {
    payload: commands.map(serialize),
    callbacks
  }
}

async function resolveHeaderBarsCapabilities(
  host: HeaderBarsHost
): Promise<HeaderBarsCapabilities> {
  const cached = host.getCapabilitiesState()
  if (cached.protocolVersion === 2) {
    return cached
  }
  try {
    const capabilities = await host.invokeHeaderBars<HeaderBarsCapabilities>(
      HEADER_BARS_METHOD.getCapabilities
    )
    host.setCapabilitiesState(capabilities)
    return capabilities
  } catch (error: unknown) {
    void error
    return cached
  }
}

async function registerCommandEventCallbacks(
  host: HeaderBarsHost,
  callbacks: Map<string, HeaderBarsEventHandler | undefined>
) {
  const tasks: Array<Promise<void>> = []
  callbacks.forEach((callback, key) => {
    host.getEventOverridesMap().set(key, callback)
    const separatorIndex = key.lastIndexOf(':')
    const id = key.slice(0, separatorIndex)
    const event = key.slice(separatorIndex + 1)
    tasks.push(
      host.invokeHeaderBars<undefined>(
        HEADER_BARS_METHOD.setCommandEventCallbackEnabled,
        {
          id,
          event,
          enabled: typeof callback === 'function'
        }
      )
    )
  })
  await Promise.all(tasks)
}

export function initHeaderBarsFacade(host: HeaderBarsHost): HeaderBarsFacade {
  const facade: HeaderBarsFacade = {
    visible: false,
    onTitleChange: undefined,
    getVisible: async () => {
      return await syncHeaderBarsVisible(host)
    },
    setVisible: async (visible: boolean) => {
      await setHeaderBarsVisible(host, visible)
    },
    addCommand: async (
      command: HeaderBarsCommandDefinition,
      posCommand: string,
      pos: 'before' | 'after' = 'after'
    ) => {
      const { onClick, ...commandPayload } = command
      const added = await host.invokeHeaderBars<boolean>(
        HEADER_BARS_METHOD.addCommand,
        { command: commandPayload, posCommand, pos }
      )
      const clickHandler = onClick
      if (added && typeof clickHandler === 'function') {
        host.getOverridesMap().set(command.id, clickHandler)
        await host.invokeHeaderBars<undefined>(
          HEADER_BARS_METHOD.setCommandCallbackEnabled,
          {
            id: command.id,
            enabled: true
          }
        )
      }
      return added
    },
    getCommand: (id: string) => getHeaderBarsCommandRef(host, id),
    listViewCommands: async () => {
      const capabilities = await resolveHeaderBarsCapabilities(host)
      const commands = await host.invokeHeaderBars<HeaderBarsCommandState[]>(
        HEADER_BARS_METHOD.listViewCommands,
        { protocolVersion: capabilities.protocolVersion }
      )
      syncHeaderBarsCommands(host, commands)
      return commands
    },
    getCapabilities: async () => {
      return await resolveHeaderBarsCapabilities(host)
    },
    addCommands: async (
      commands: HeaderBarsCommandDefinition[],
      posCommand: string,
      pos: 'before' | 'after' = 'after'
    ) => {
      const capabilities = await resolveHeaderBarsCapabilities(host)
      if (
        capabilities.protocolVersion !== 2 ||
        !capabilities.features.treeCommands ||
        !capabilities.features.batchCommands
      ) {
        return {
          success: false,
          code: 'HEADER_BARS_PROTOCOL_UNSUPPORTED',
          message: 'HeaderBars extension is not supported'
        }
      }
      const { payload, callbacks } = serializeCommandDefinitions(commands)
      try {
        const result = await host.invokeHeaderBars<HeaderBarsMutationResult>(
          HEADER_BARS_METHOD.addCommands,
          { commands: payload, posCommand, pos }
        )
        if (!result.success) {
          return result
        }
        await registerCommandEventCallbacks(host, callbacks)
        return result
      } catch (error: unknown) {
        return {
          success: false,
          code: 'HEADER_BARS_TRANSPORT_ERROR',
          message: error instanceof Error ? error.message : String(error)
        }
      }
    },
    setSectionVisible: async (section, visible) => {
      const capabilities = await resolveHeaderBarsCapabilities(host)
      if (
        capabilities.protocolVersion !== 2 ||
        !capabilities.features.sectionVisibility
      ) {
        return {
          success: false,
          code:
            capabilities.protocolVersion === 2
              ? 'HEADER_BARS_VISIBILITY_NOT_SUPPORTED'
              : 'HEADER_BARS_PROTOCOL_UNSUPPORTED',
          message:
            capabilities.protocolVersion === 2
              ? 'HeaderBars section visibility is not supported'
              : 'HeaderBars extension is not supported'
        }
      }
      try {
        return await host.invokeHeaderBars<HeaderBarsMutationResult>(
          HEADER_BARS_METHOD.setSectionVisible,
          { section, visible }
        )
      } catch (error: unknown) {
        return {
          success: false,
          code: 'HEADER_BARS_TRANSPORT_ERROR',
          message: error instanceof Error ? error.message : String(error)
        }
      }
    }
  }

  Object.defineProperty(facade, 'visible', {
    configurable: true,
    enumerable: true,
    get: () => host.getVisibleState(),
    set: (next: boolean) => {
      setHeaderBarsVisible(host, next).catch((err: unknown) => {
        host.emitHeaderBarsError('set headerBars.visible failed', err)
      })
    }
  })

  Object.defineProperty(facade, 'onTitleChange', {
    configurable: true,
    enumerable: true,
    get: () => host.getTitleHandler(),
    set: (handler: HeaderBarsTitleChangeHandler | undefined) => {
      host.setTitleHandler(handler)
      if (typeof handler !== 'function') {
        return
      }
      ensureHeaderBarsTitleChangeSubscription(host).catch((err: unknown) => {
        host.emitHeaderBarsError('subscribe headerBars titleChange failed', err)
      })
    }
  })

  return facade
}

export async function ensureHeaderBarsTitleChangeSubscription(
  host: HeaderBarsHost
): Promise<void> {
  if (host.isTitleSubscribed()) {
    return
  }
  host.setTitleSubscribed(true)
  host.onInternalTitleChange((title: unknown) => {
    if (typeof title !== 'string') {
      return
    }
    host.getTitleHandler()?.(title)
  })
  try {
    await host.subscribeEditorTitleChange()
  } catch (error: unknown) {
    host.setTitleSubscribed(false)
    throw error
  }
}

export function syncHeaderBarsCommands(
  host: HeaderBarsHost,
  commands: HeaderBarsCommandState[]
) {
  const commandMap = host.getCommandsMap()
  commandMap.clear()
  const visit = (command: HeaderBarsCommandState) => {
    commandMap.set(command.id, command)
    command.subItems?.forEach((item) => visit(item as HeaderBarsCommandState))
  }
  for (const command of commands) {
    visit(command)
  }
}

export function applyHeaderBarsChanged(
  host: HeaderBarsHost,
  payload?: HeaderBarsChangedPayload
) {
  const snapshot = payload?.snapshot
  if (!snapshot) {
    return
  }
  host.setVisibleState(snapshot.visible)
  syncHeaderBarsCommands(host, snapshot.commands)
}

export async function syncHeaderBarsVisible(
  host: HeaderBarsHost
): Promise<boolean> {
  const payload = await host.invokeHeaderBars<{ visible: boolean }>(
    HEADER_BARS_METHOD.getVisible
  )
  host.setVisibleState(payload.visible)
  return host.getVisibleState()
}

export async function setHeaderBarsVisible(
  host: HeaderBarsHost,
  visible: boolean
) {
  host.setVisibleState(visible)
  await host.invokeHeaderBars<undefined>(HEADER_BARS_METHOD.setVisible, {
    visible
  })
}

export function getHeaderBarsCommandRef(
  host: HeaderBarsHost,
  id: string
): HeaderBarsCommandRef {
  const refs = host.getRefsMap()
  const existing = refs.get(id)
  if (existing) {
    return existing
  }

  const commands = host.getCommandsMap()
  if (!commands.has(id)) {
    resolveHeaderBarsCapabilities(host)
      .then(async (capabilities) => {
        return await host.invokeHeaderBars<{
          command: HeaderBarsCommandState | null
        }>(HEADER_BARS_METHOD.getCommand, {
          id,
          protocolVersion: capabilities.protocolVersion
        })
      })
      .then((payload) => {
        if (payload.command) {
          commands.set(payload.command.id, payload.command)
        }
      })
      .catch((err: unknown) => {
        host.emitHeaderBarsError('fetch headerBars command failed', err)
      })
  }

  const ref: HeaderBarsCommandRef = {
    id,
    visible: true,
    disabled: false,
    active: false,
    src: undefined,
    label: undefined,
    editable: undefined,
    open: false,
    style: undefined,
    subItems: undefined,
    onCommandClick: undefined,
    onCommandOpen: undefined,
    onCommandClose: undefined,
    getState: () => commands.get(id)
  }

  Object.defineProperties(ref, {
    visible: {
      configurable: true,
      enumerable: true,
      get: () => commands.get(id)?.visible !== false,
      set: (next: boolean) => {
        const current = commands.get(id)
        if (current) {
          commands.set(id, { ...current, visible: next })
        }
        host
          .invokeHeaderBars<undefined>(HEADER_BARS_METHOD.setCommandVisible, {
            id,
            visible: next
          })
          .catch((err: unknown) => {
            host.emitHeaderBarsError(
              'set headerBars command visible failed',
              err
            )
          })
      }
    },
    disabled: {
      configurable: true,
      enumerable: true,
      get: () => commands.get(id)?.disabled === true,
      set: (next: boolean) => {
        const current = commands.get(id)
        if (current) {
          commands.set(id, { ...current, disabled: next })
        }
        host
          .invokeHeaderBars<undefined>(HEADER_BARS_METHOD.setCommandDisabled, {
            id,
            disabled: next
          })
          .catch((err: unknown) => {
            host.emitHeaderBarsError(
              'set headerBars command disabled failed',
              err
            )
          })
      }
    },
    active: {
      configurable: true,
      enumerable: true,
      get: () => commands.get(id)?.active === true,
      set: (next: boolean) => {
        const current = commands.get(id)
        if (current) {
          commands.set(id, { ...current, active: next })
        }
        host
          .invokeHeaderBars<undefined>(HEADER_BARS_METHOD.setCommandActive, {
            id,
            active: next
          })
          .catch((err: unknown) => {
            host.emitHeaderBarsError(
              'set headerBars command active failed',
              err
            )
          })
      }
    },
    src: {
      configurable: true,
      enumerable: true,
      get: () => commands.get(id)?.src,
      set: (next: string | undefined) => {
        if (typeof next !== 'string') {
          return
        }
        const current = commands.get(id)
        if (current) {
          commands.set(id, { ...current, src: next })
        }
        host
          .invokeHeaderBars<undefined>(HEADER_BARS_METHOD.setCommandSrc, {
            id,
            src: next
          })
          .catch((err: unknown) => {
            host.emitHeaderBarsError('set headerBars command src failed', err)
          })
      }
    },
    label: {
      configurable: true,
      enumerable: true,
      get: () => commands.get(id)?.label,
      set: (next: string | undefined) => {
        if (
          typeof next !== 'string' ||
          id === 'title' ||
          id === 'save-status'
        ) {
          return
        }
        const current = commands.get(id)
        if (current) {
          commands.set(id, { ...current, label: next })
        }
        host
          .invokeHeaderBars<undefined>(HEADER_BARS_METHOD.setCommandLabel, {
            id,
            label: next
          })
          .catch((err: unknown) => {
            host.emitHeaderBarsError('set headerBars command label failed', err)
          })
      }
    },
    editable: {
      configurable: true,
      enumerable: true,
      get: () => commands.get(id)?.editable,
      set: (next: boolean | undefined) => {
        if (id !== 'title') {
          host.emitHeaderBarsError(
            'headerBars command editable is only supported for title',
            new Error('headerBars command editable is only supported for title')
          )
          return
        }
        const current = commands.get(id)
        if (current) {
          commands.set(id, { ...current, editable: next })
        }
        host
          .invokeHeaderBars<undefined>(HEADER_BARS_METHOD.setCommandEditable, {
            id,
            editable: next
          })
          .catch((err: unknown) => {
            host.emitHeaderBarsError(
              'set headerBars command editable failed',
              err
            )
          })
      }
    },
    open: {
      configurable: true,
      enumerable: true,
      get: () => commands.get(id)?.open === true,
      set: (next: boolean) => {
        host
          .invokeHeaderBars<HeaderBarsMutationResult>(
            HEADER_BARS_METHOD.setCommandOpen,
            { id, open: next }
          )
          .then((result) => {
            if (!result.success) {
              host.emitHeaderBarsError(
                'set headerBars command open failed',
                new Error(result.message ?? result.code ?? 'unknown error')
              )
            }
          })
          .catch((err: unknown) => {
            host.emitHeaderBarsError('set headerBars command open failed', err)
          })
      }
    },
    style: {
      configurable: true,
      enumerable: true,
      get: () => commands.get(id)?.style,
      set: (next: HeaderBarsCommandStyle | undefined) => {
        const current = commands.get(id)
        if (current) {
          commands.set(id, { ...current, style: next })
        }
        host
          .invokeHeaderBars<HeaderBarsMutationResult>(
            HEADER_BARS_METHOD.setCommandOptions,
            { id, options: { style: next } }
          )
          .then((result) => {
            if (!result.success) {
              host.emitHeaderBarsError(
                'set headerBars command style failed',
                new Error(result.message ?? result.code ?? 'unknown error')
              )
            }
          })
          .catch((err: unknown) => {
            host.emitHeaderBarsError('set headerBars command style failed', err)
          })
      }
    },
    subItems: {
      configurable: true,
      enumerable: true,
      get: () => commands.get(id)?.subItems,
      set: (next: HeaderBarsCommandDefinition[] | undefined) => {
        const current = commands.get(id)
        if (current) {
          commands.set(id, { ...current, subItems: next })
        }
        host
          .invokeHeaderBars<HeaderBarsMutationResult>(
            HEADER_BARS_METHOD.setCommandOptions,
            { id, options: { subItems: next } }
          )
          .then((result) => {
            if (!result.success) {
              host.emitHeaderBarsError(
                'set headerBars command subItems failed',
                new Error(result.message ?? result.code ?? 'unknown error')
              )
            }
          })
          .catch((err: unknown) => {
            host.emitHeaderBarsError(
              'set headerBars command subItems failed',
              err
            )
          })
      }
    },
    onCommandClick: {
      configurable: true,
      enumerable: true,
      get: () => host.getOverridesMap().get(id),
      set: (handler: (() => void | Promise<void>) | undefined) => {
        host.getOverridesMap().set(id, handler)
        host
          .invokeHeaderBars<undefined>(
            HEADER_BARS_METHOD.setCommandCallbackEnabled,
            {
              id,
              enabled: typeof handler === 'function'
            }
          )
          .catch((err: unknown) => {
            host.emitHeaderBarsError(
              'set headerBars command callback failed',
              err
            )
          })
      }
    },
    onCommandOpen: {
      configurable: true,
      enumerable: true,
      get: () =>
        host.getEventOverridesMap().get(`${id}:open`) as
          | HeaderBarsCommandOpenHandler
          | undefined,
      set: (handler: HeaderBarsCommandOpenHandler | undefined) => {
        host.getEventOverridesMap().set(`${id}:open`, handler)
        host
          .invokeHeaderBars<undefined>(
            HEADER_BARS_METHOD.setCommandEventCallbackEnabled,
            { id, event: 'open', enabled: typeof handler === 'function' }
          )
          .catch((err: unknown) => {
            host.emitHeaderBarsError(
              'set headerBars command open callback failed',
              err
            )
          })
      }
    },
    onCommandClose: {
      configurable: true,
      enumerable: true,
      get: () =>
        host.getEventOverridesMap().get(`${id}:close`) as
          | HeaderBarsCommandCloseHandler
          | undefined,
      set: (handler: HeaderBarsCommandCloseHandler | undefined) => {
        host.getEventOverridesMap().set(`${id}:close`, handler)
        host
          .invokeHeaderBars<undefined>(
            HEADER_BARS_METHOD.setCommandEventCallbackEnabled,
            { id, event: 'close', enabled: typeof handler === 'function' }
          )
          .catch((err: unknown) => {
            host.emitHeaderBarsError(
              'set headerBars command close callback failed',
              err
            )
          })
      }
    }
  })

  refs.set(id, ref)
  return ref
}

export async function subscribeHeaderBarsTitleChange(
  invoke: <T>(method: string, args: unknown[]) => Promise<T>
) {
  await invoke(InvokeMethod.ListenEditorEvent, ['titleChange'])
}
