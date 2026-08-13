### 安装

```bash
npm install --save weboffice-js-sdk
```

### 基本用法

#### npm 包

```javascript
import { connect } from 'weboffice-js-sdk'

const sdk = await connect({
  fileId: '您系统中的 file id',
  endpoint: '石墨服务的地址',
  signature: '用您的 app id 和 secret 签发的签名',
  token: '用于您系统识别用户请求的 token',
  container: document.querySelector('#shimo-file'), // iframe 挂载的目标容器元素
  lang: 'en-US', // 可选；未指定时使用编辑器默认语言
  headerBarsVisible: true, // 顶部栏初始是否展示, 默认值 true; 传入 false 表示隐藏
  userUuid: '您的 uuid' // 仅在 v2 版本回调时需要传入（co-1.3+ 支持）
})

// 获取编辑器实例
const editor = sdk.getEditor()

// 监听保存状态变化
editor.on('saveStatusChanged', (payload) => {
  console.log('保存状态:', payload.status)
})
```

---

#### 使用`<script>`

1. 使用 [npm view](https://docs.npmjs.com/cli/v7/commands/npm-view) 和 [npm pack](https://docs.npmjs.com/cli/v7/commands/npm-pack) 下载代码包 (`.tgz` 格式)
2. 将 `.tgz` 解压缩后的 `dist` 目录下的文件放置到您托管静态资源的空间，然后使用 `<script>` 引入 `index.js` 资源
3. 通过 `window.WebOfficeJSSDK` 对象获取对应的方法

```js
const { connect, FileType } = window.WebOfficeJSSDK
// 等价于
const { connect, FileType } = require('weboffice-js-sdk')
```

#### 使用示例

```js
const { connect } = require('weboffice-js-sdk')

const fileId = '1234'
const uuid = 'youruuid'

// 从您的后端服务获取用于石墨鉴权的签名和 token
const { signature, token } = await getCredentialsFromServer()

connect({
  fileId: fileId,
  endpoint: 'https://shimo-sdk-endpoint/', // endpoint 因环境而异，请联系技术支持
  signature: signature,
  token: token,
  container: document.querySelector('#shimo-file'), // iframe 挂载的目标容器元素
  userUuid: uuid
}).then((sdk) => {
  // sdk 即为 OfficeSDK 实例
})
```

调用 `connect()` 时，会以传入参数为基础，初始化一个 `<iframe>` 并插入 `container` 对应的元素中。

返回的 `sdk` 为 `OfficeSDK` 实例，用于和 SDK、编辑器交互。

### 常用 connectOptions 配置

#### 凭证自动刷新（建议）

`signature` 和 `token` 存在有效期。建议在签名过期前自动刷新，避免用户长时间编辑时凭证失效：

```typescript
const { expireMs } = (await appService.getExpireConfig()).data

const options: ConnectOptions = {
  ...config,
  refreshCredentialsInterval: Math.ceil(expireMs * 0.8),
  getCredentials: async () => (await appService.getCredentials()).data
}

const sdk = await connect(options)
```

示例中，`GET /api/apps/expire-config` 返回过期时长 `expireMs`，`GET /api/credentials` 返回新的 `signature` 和 `token`。接口路径可按接入方系统实际情况调整。

#### headerBarsVisible

通过 `headerBarsVisible` 控制顶部栏初始是否展示：

```js
const officeSDK = await connect({
  ...options,
  headerBarsVisible: false
})
```

- `true` 或不传：初始展示顶部栏
- `false`：初始隐藏顶部栏
- 连接后如需动态切换，使用 `await officeSDK.headerBars.setVisible(visible)`

#### 国际化：编辑器多语言（可选，co-1.8+）

调用 `connect` 时可通过 `lang` 指定编辑器界面语言；用户选择“系统默认”时省略 `lang`（或传 `undefined`；纯 JavaScript 也可传 `null`）：

```typescript
const options: ConnectOptions = {
  ...config,
  ...(editorLang ? { lang: editorLang } : {}) // “系统默认”时不传 lang
}

const sdk = await connect(options)
```

1.8 版本编辑器支持以下 16 种语言：

| 语言          | `lang` 语言码 | 语言             | `lang` 语言码 |
| ------------- | ------------- | ---------------- | ------------- |
| 简体中文      | `zh-CN`       | 繁體中文         | `zh-TW`       |
| English       | `en-US`       | 日本語           | `ja-JP`       |
| 한국어        | `ko-KR`       | Español          | `es-ES`       |
| Português     | `pt-PT`       | Deutsch          | `de-DE`       |
| Français      | `fr-FR`       | Italiano         | `it-IT`       |
| Русский       | `ru-RU`       | Bahasa Indonesia | `id-ID`       |
| Tiếng Việt    | `vi-VN`       | ไทย              | `th-TH`       |
| Bahasa Melayu | `ms-MY`       | العربية          | `ar-SA`       |

为兼容旧版写法，仍可传入 `en`、`ja`，SDK 会分别映射为 `en-US`、`ja-JP`。新接入建议使用表中的标准语言码。

### 如何处理 URL

由于石墨 SDK 以 `iframe` 的形式挂载到当前页面，`iframe.src` 对应的 URL 并不适合用于分享，而且在一些功能上，比如 @ 文件，需要用到您系统中对应的 URL 格式，比如 `https://your-domain/files/:id`。

为了解决这个问题，石墨 SDK 引入 `generateUrl()` 和 `openLink()` 方法：

```js
import { UrlSharingType } from 'weboffice-js-sdk'

const officeSDK = await connect({
  ...,

  generateUrl(fileId: string, info: GenerateUrlInfo): string {
    if (info?.sharingType === UrlSharingType.FormFill) {
      return `https://your-domain/files/${fileId}/fill-form`
    }

    if (info?.sharingText) {
      return `https://your-domain/files/${fileId} ${info.sharingText}`
    }

    return `https://your-domain/files/${fileId}`
  },

  openLink(url: string): void {
    // 以 React Router 为例

    // 假设 url 是 'https://your-domain/files/1'，在当前页跳转，其他则新窗口打开
    if (url.includes('your-domain/files/')) {
      const u = new URL(url)
      history.push(u.pathname)
    } else {
      window.open(url)
    }
  },

  // 从当前 url 中解析出文件 id 并返回
  // 假设 url 是 'https://your-domain/files/123'，则返回 { fileId: '123' }
  getFileInfoFromUrl(url: string): {fileId:string} {
    let fromId
    const urlWithoutParams = url.split('?')[0]
    let splitPath = urlWithoutParams.split('/')
    fromId = splitPath[splitPath.length - 1]
    return Promise.resolve({
        fileId: fromId
    })
  }
})
```

#### URL 的上下文信息

为了在 URL 上传递上下文信息，比如 URL 指向的段落、单元格，在调用 `generateUrl()` 生成 URL 后，会在 URL 后附加一个 `smParams=PARAMS` 的参数：

```
https://your-domain/files/:id?smParams=PARAMS
```

**如无特殊需要，请保留该参数。**

默认情况下，调用 `connect()` 会从当前 `location.search` 中提取 `smParams`，如果遇到需要自定义参数的场合，可以通过 `connect({ smParams: PARAMS })` 参数修改。

`smParams` 为经过 [base62str](https://www.npmjs.com/package/base62str) 序列化后的 `Record<string, unknown>` 对象。

前端可以使用 `base62str` 生成 `smParams`：

```js
const Base62Str = require('base62str').default
const base62 = Base62Str.createInstance()

const obj = {
  type: 'form',
  path: '%2Ffill'
}

// 固定字段顺序，保证相同数据生成稳定的编码结果
const json = JSON.stringify(obj, ['type', 'fileGuid', 'path'])
console.log('JSON:', json)

const smParams = base62.encodeStr(json)
console.log('encodeStr:', smParams)

// 部分版本也提供 encode() 别名
if (typeof base62.encode === 'function') {
  console.log('encode:', base62.encode(json))
}

connect({
  smParams
})
```

**在传入 `smParams` 参数时，将不会从 `location.search` 中获取数据**，如果想保留原有信息，可以这样传递：

```js
const paramsList: Array<string | Record<string, unknown>>

const originParams = new URLSearchParams(location.search).get('smParams')
// 保留原来的上下文信息
if (originParams) {
  paramsList.push(originParams)
}

// 添加自定义的上下文信息
paramsList.push({
  myVar: 'myVal'
})

connect({
  smParams: paramsList
})
```

#### URL Info

`generateUrl(fileId, info)` 中的 `info` 是用于对 URL 进行一些特殊处理的。

`sharingText`：石墨默认提供的分享文本：比如

- `https://your-domain/files/1 xxx 邀请您参与《标题》协作，请复制粘贴后在浏览器打开`
- `https://your-domain/files/1/fill-form xxx 邀请您填写《标题》表单，……`

`sharingType`：表示此次 `generateUrl()` 对应的行为类型，比如：

- `UrlSharingType.Form` 代表一般的打开编辑表单的行为
- `UrlSharingType.FormPreview` 代表打开预览表单页面的行为
- `UrlSharingType.FormFill` 代表打开填写表单页面的行为

您需要根据具体类型，生成不同的 URL，比如：

- `UrlSharingType.Form`、`UrlSharingType.FormPreview` 等一般需要进行鉴权，因此可以用 `/files/${fileId}`
- `UrlSharingType.FormFill` 填写表单一般不需要登录鉴权，因此可以用另一个独立的路由，比如 `/files/${fileId}/fill-form`

在实际操作中，您可以根据 `sharingType` 按需为 URL 添加分享文本。**若添加了分享文本，则需要您在 `parseUrl()` 中对 URL 进行处理**，比如：

```js
// url: 'https://your-domain/files/1 xxx 邀请您参与《标题》协作，请复制粘贴后在浏览器打开'
parseUrl(url: string) {
  return url.split(' ')[0] // 返回 'https://your-domain/files/1
}
```

#### 打开表格指定工作表 (Sheet)

**使用本章节用法时，请先了解 [URL 的上下文信息](#url-的上下文信息) 章节**。

此用法适用于表格中存在多个工作表 (Sheet) ，希望在打开编辑器时，直接展示某个工作表格而非默认的第一个工作表。如用于希望直接分享表格的某个工作表链接给其他协作者，他人在打开后可直接查看指定的工作表。

首先通过表格编辑器接口 [sdk.workbook](https://support.shimo.net/apidoc/docs-site/6000010/doc-338262#sdkworkbook) 的 `getActiveWorksheet()` 获取当前激活工作表，并读取其 `id`。此 ID 可追加在接入方自身的 URL 上作为参数。

如通过 `URL QueryString` 方式传递：`https://your-domain.com/files/abcdefg?sheetId=XXXXX&smParams=XXXXXXXXXXXXXXXXXXXXXX`

`sheetId` 仅为参数名举例，接入方可结合自身业务命名。

```js
const paramsList: Array<string | Record<string, unknown>>
const queryParams = new URLSearchParams(location.search)

const originParams = queryParams.get('smParams')
const sheetId = queryParams.get('sheetId')

// 保留原来的上下文信息
if (originParams) {
  paramsList.push(originParams)
}
// paramsList
// => [originParamsStringValue]

// 添加自定义的上下文信息
paramsList.push({ sheetId: '通过 QueryString 中获取的 sheetId' })
// paramsList
// => [originParamsStringValue, {"sheetId": "XXXXX"}]

connect({
  smParams: paramsList
})
```

#### 定位文中位置

说明：打开编辑器时，定位至在正文中 at 某用户或评论的位置

支持类型：

- `文档` - `document`
- `表格` - `spreadsheet`
- `文稿` - `documentPro`

**使用本章节用法时，请先了解 [URL 的上下文信息](#url-的上下文信息) 章节**。

此用法适用于:

- 定位@用户： 在接入方系统的文件中 at 了指定用户，在回调接口中收到 `石墨 SDK 事件` 中的 `mention_at` 类型事件，并获取 `mentionAt.guid` 字段作为参数拼接至接入方的访问链接上，在接入方系统通知对应用户时，推送的链接可直接打开对应文件并定位至当前用户被 at 的正文位置，以便于查看对应位置相关内容。
- 新增评论： 在接入方系统的文件中新增了评论，在回调接口中收到 `石墨 SDK 事件` 中的 `comment` 类型事件，并获取 `comment.selectionGuid` 字段作为参数拼接至接入方的访问链接上，在接入方系统通知对应用户时，推送的链接可直接打开对应文件并定位至当前新增的评论位置，以便于查看对应位置相关内容。

如通过 `URL QueryString` 方式传递：`https://your-domain.com/files/abcdefg?mentionId=XXXXX&smParams=XXXXXXXXXXXXXXXXXXXXXX`

`mentionId` 仅为参数名举例，接入方可结合自身业务命名。

```js
const paramsList: Array<string | Record<string, unknown>>
const queryParams = new URLSearchParams(location.search)

const originParams = queryParams.get('smParams')
const mentionId = queryParams.get('mentionId')

// 保留原来的上下文信息
if (originParams) {
  paramsList.push(originParams)
}
// paramsList
// => [originParamsStringValue]

// 添加自定义的上下文信息
paramsList.push({ hash: '通过 QueryString 中获取的 mentionId' })
// paramsList
// => [originParamsStringValue, {"hash": "XXXXX"}]

connect({
  smParams: paramsList
})
```

## 支持的文档类型

| 类型         | 说明     | 模块                                            |
| ------------ | -------- | ----------------------------------------------- |
| Document     | 文档     | [Document](./suite/document.md)                 |
| DocumentPro  | 文稿     | [DocumentPro](./suite/document-pro.md)          |
| Spreadsheet  | 表格     | [Spreadsheet](../docs/modules/Spreadsheet.md)   |
| Table        | 简单表格 | [Table](../docs/modules/Table.md)               |
| Presentation | 演示文稿 | [Presentation](../docs/modules/Presentation.md) |
| Flowchart    | 流程图   | [Flowchart](../docs/modules/Flowchart.md)       |
| Form         | 表单     | [Form](../docs/modules/Form.md)                 |

## 获取帮助

如果您在使用过程中遇到问题，请参考：

- [详细 API 文档](https://github.com/shimo-open/weboffice-js-sdk/blob/master/docs/modules) - 详细的接口文档
- [石墨 SDK 2.0 官网](https://open.shimo.im/) - 产品详细介绍

## 注意事项

> ⚠️ 此 SDK 仅适用于石墨 SDK 2.0 产品，无法用于石墨文档官网产品 (shimo.im)。
