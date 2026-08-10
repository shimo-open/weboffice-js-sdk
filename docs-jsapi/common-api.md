# 公共处理方法

## 概述

本页描述 `OfficeSDK` 实例上的公共 API，也就是不依赖具体编辑器套件、宿主侧可以统一调用的能力。

这类能力主要用于：

- 建立连接并等待 SDK 就绪
- 读取当前文件类型和加载状态
- 监听 SDK / 编辑器通用事件
- 获取当前编辑器实例
- 断开连接
- 更新鉴权信息

`HeaderBars` 虽然也挂在 `sdk` 根级实例上，但它属于独立扩展能力，不在本页展开。请查看 [HeaderBars](./headerbars.md)。

---

## 使用示例

```typescript
import { connect, Event, FileType } from 'weboffice-js-sdk'

const sdk = await connect(options)

await sdk.ready()

sdk.on(Event.Error, (error) => {
  console.error(error)
})

if (sdk.fileType === FileType.Spreadsheet) {
  const editor = sdk.getEditor()
  await editor.showHistory?.()
}
```

---

## 方法列表

| API                           | 说明                                  |
| ----------------------------- | ------------------------------------- |
| `connect(options)`            | 创建并连接 SDK，返回 `OfficeSDK` 实例 |
| `sdk.ready()`                 | 等待 SDK 进入可交互状态               |
| `sdk.on(event, listener)`     | 监听事件                              |
| `sdk.once(event, listener)`   | 监听一次性事件                        |
| `sdk.off(event, listener?)`   | 取消事件监听                          |
| `sdk.fileType`                | 获取当前文件类型                      |
| `sdk.readyState`              | 获取当前 SDK 状态                     |
| `sdk.getEditor()`             | 获取当前编辑器实例                    |
| `sdk.disconnect()`            | 断开连接并移除 iframe                 |
| `sdk.setCredentials(payload)` | 更新鉴权信息                          |
| `sdk.setSignature(signature)` | 仅更新 signature，已废弃              |
| `sdk.setToken(token)`         | 仅更新 token，已废弃                  |
| `sdk.getPerformanceEntries()` | 获取性能信息片段列表                  |

---

## connect(options)

### 说明

创建并初始化一个 `OfficeSDK` 实例。

```typescript
import { connect } from 'weboffice-js-sdk'

const sdk = await connect({
  fileId: 'your-file-id',
  endpoint: 'https://your-shimo-endpoint',
  signature: 'your-signature',
  token: 'your-token',
  container: document.querySelector('#shimo-file'),
  headerBarsVisible: true,
  lang: 'zh-CN'
})
```

### 常用配置

| 配置项                       | 说明                                                                         |
| ---------------------------- | ---------------------------------------------------------------------------- |
| `fileId`                     | 要打开的文档 ID                                                              |
| `endpoint`                   | 石墨 SDK 服务地址                                                            |
| `signature`                  | 用于石墨 SDK 鉴权的签名                                                      |
| `token`                      | 用于接入方系统鉴权的 token                                                   |
| `container`                  | iframe 挂载的目标容器元素                                                    |
| `refreshCredentialsInterval` | 自动刷新鉴权信息的时间间隔，单位为毫秒                                       |
| `getCredentials`             | 获取最新 `signature` 和 `token` 的异步方法                                   |
| `headerBarsVisible`          | 顶部栏初始是否展示，`false` 表示隐藏                                         |
| `lang`                       | 编辑器 locale，例如 `zh-CN`、`en-US`；未传时使用 iframe 服务端设置的默认语言 |
| `smParams`                   | URL 上下文参数，可传 base62 字符串、对象，或由字符串和对象组成的数组         |

### 自动刷新鉴权

`signature` 和 `token` 存在有效期。建议通过 `refreshCredentialsInterval` 和 `getCredentials` 配置自动刷新，避免用户长时间停留页面后因鉴权过期而中断编辑。

```typescript
const officeSDK = await connect({
  signature: '[your signature]',
  token: '[your token]',
  // 更新鉴权的时间间隔，单位为毫秒
  // 若过期时间为 7 天，则建议设置为 3.5 天
  refreshCredentialsInterval: 1000 * 3600 * 24 * 3.5,
  getCredentials: async () => {
    const res = await getCredentialsFromServer()
    return {
      signature: res.signature,
      token: res.token
    }
  }
})
```

`getCredentials` 应从接入方后端获取最新鉴权信息，并返回同时包含 `signature` 和 `token` 的对象。

### 顶部栏初始显示状态

通过 `headerBarsVisible` 控制顶部栏初始是否展示：

```typescript
const sdk = await connect({
  ...options,
  headerBarsVisible: false
})
```

- `true` 或不传：初始展示顶部栏
- `false`：初始隐藏顶部栏
- 连接后如需动态切换，使用 `await sdk.headerBars.setVisible(visible)`，详见 [HeaderBars](./headerbars.md)

### locale（编辑器语言）

通过 `lang` 指定编辑器 locale：

```typescript
const sdk = await connect({
  ...options,
  lang: 'en-US'
})
```

支持的标准值：

- `zh-CN`、`zh-TW`
- `en-US`、`ja-JP`、`ko-KR`
- `es-ES`、`pt-PT`、`de-DE`、`fr-FR`、`it-IT`、`ru-RU`
- `id-ID`、`vi-VN`、`th-TH`、`ms-MY`、`ar-SA`

为兼容旧版写法，仍可传入 `en`、`ja`，SDK 会分别映射为 `en-US`、`ja-JP`。未传 `lang` 时，iframe 使用服务端设置的默认语言。

### 返回值

返回 `Promise<OfficeSDK>`。

### 说明补充

- `connect()` 成功后返回 `sdk`
- 后续所有公共 API 和编辑器 API 都从该实例继续调用

---

## sdk.ready()

### 说明

等待 SDK 进入 `Ready` 状态。若当前已经 ready，则会直接返回。

```typescript
const sdk = await connect(options)

await sdk.ready()
```

### 返回值

返回 `Promise<void>`。

### 适用场景

- 在调用具体编辑器能力前等待可交互状态
- 需要确保 `fileType`、根级 API 或编辑器方法已经可用

---

## sdk.on(event, listener)

### 说明

监听 SDK 或编辑器转发出来的通用事件。

```typescript
import { Event } from 'weboffice-js-sdk'

sdk.on(Event.Error, (error) => {
  console.error(error)
})

sdk.on(Event.ReadyState, (payload) => {
  console.log(payload.state)
})
```

### 常见事件

- `Event.Error`：错误事件
- `Event.ReadyState`：SDK 状态变化事件
- `Event.EditorRendered`：编辑器完成首屏渲染

### 返回值

返回当前 `sdk` 实例，便于链式调用。

---

## sdk.once(event, listener)

### 说明

监听一次性事件，回调触发后会自动取消监听。

```typescript
import { Event } from 'weboffice-js-sdk'

sdk.once(Event.EditorRendered, () => {
  console.log('editor first render finished')
})
```

### 返回值

返回当前 `sdk` 实例。

---

## sdk.off(event, listener?)

### 说明

取消事件监听。

```typescript
import { Event } from 'weboffice-js-sdk'

const handleError = (error: Error) => {
  console.error(error)
}

sdk.on(Event.Error, handleError)
sdk.off(Event.Error, handleError)
```

### 参数

- `event`：事件名
- `listener`：要移除的监听函数；不传时会移除该事件下的监听

### 返回值

返回当前 `sdk` 实例。

---

## sdk.fileType

### 说明

读取当前文件类型。

```typescript
import { FileType } from 'weboffice-js-sdk'

if (sdk.fileType === FileType.Document) {
  console.log('当前是轻文档')
}
```

### 类型

```typescript
sdk.fileType: FileType
```

### 适用场景

- 根据当前套件决定后续调用哪组编辑器 API
- 在一个统一入口中分发不同文档类型的处理逻辑

---

## sdk.readyState

### 说明

读取当前 SDK 状态。

```typescript
console.log(sdk.readyState)
```

### 类型

```typescript
sdk.readyState: ReadyState
```

### 说明补充

常见状态包括：

- `Loading`
- `LoadingEditor`
- `Ready`
- `Failed`

---

## sdk.getEditor()

### 说明

获取当前文件对应的编辑器实例。

```typescript
const editor = sdk.getEditor()

await editor.showHistory?.()
```

### 返回值

返回当前套件对应的编辑器实例类型。

### 说明补充

- 这是兼容历史代码的通用入口
- 如果当前文档已经提供根级 facade，新增代码优先使用 API 写法

---

## sdk.disconnect()

### 说明

断开 SDK 连接，并从宿主容器中移除当前 iframe。

```typescript
sdk.disconnect()
```

### 适用场景

- 页面卸载
- 路由切换
- 需要主动销毁当前编辑器实例

---

## sdk.setCredentials(payload)

### 说明

实时更新鉴权 `signature` 和 `token`。

```typescript
await sdk.setCredentials({
  signature: 'next-signature',
  token: 'next-token'
})
```

### 参数

- `payload.signature`：新的石墨鉴权签名
- `payload.token`：新的业务 token

### 返回值

返回 `Promise<void>`。

### 说明补充

- 适用于长时间停留页面后刷新鉴权信息
- 相比单独更新 `signature` 或 `token`，优先使用这个方法

---

## sdk.setSignature(signature)

### 说明

仅更新石墨鉴权 `signature`。

```typescript
await sdk.setSignature('next-signature')
```

### 返回值

返回 `Promise<void>`。

### 说明补充

该方法已废弃，建议改用 `sdk.setCredentials()`。

---

## sdk.setToken(token)

### 说明

仅更新业务 `token`。

```typescript
await sdk.setToken('next-token')
```

### 返回值

返回 `Promise<void>`。

### 说明补充

该方法已废弃，建议改用 `sdk.setCredentials()`。

---

## sdk.getPerformanceEntries()

### 说明

获取性能信息片段列表。

```typescript
const entries = await sdk.getPerformanceEntries()
console.log(entries)
```

### 返回值

返回 `Promise<PerformanceEntry[]>`。

### 说明补充

- 该列表是分段、异步采集的
- 不同时间调用，返回结果可能不完全一致
