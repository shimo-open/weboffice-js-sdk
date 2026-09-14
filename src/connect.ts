import { OfficeSDK, OfficeSDKOptions } from './OfficeSDK'
import {
  validateConnectV2Options,
  type ContentVersion,
  type ConnectV2FileType
} from './types/ConnectV2'

export {
  CONNECT_V2_FILE_TYPES,
  isConnectV2FileType,
  isConnectV2Mode,
  validateConnectV2Options
} from './types/ConnectV2'
export type { ConnectV2FileType } from './types/ConnectV2'

export interface ConnectOptions extends OfficeSDKOptions {}

export interface ConnectV2Options extends ConnectOptions {
  /**
   * V2 套件类型。当前用于 iframe 侧资源选择预留。
   */
  type: ConnectV2FileType

  /**
   * V2 打开模式。
   */
  mode: 'edit' | 'preview'
}

async function connectInternal(
  options: ConnectOptions & { type?: ConnectV2FileType },
  contentVersion?: ContentVersion
): Promise<OfficeSDK> {
  let sdk: OfficeSDK | undefined
  try {
    sdk = new OfficeSDK(options, contentVersion)
    await sdk.init()
    return sdk
  } catch (e) {
    if (options.type) {
      sdk?.disconnect()
    }
    console.log('Failed to init OfficeSDK', {
      error: e,
      options
    })
    throw e
  }
}

/**
 * 初始化 SDK，返回 Promise，当 ReadState 变为 Ready 或 Failed 时，Promise 将被 resolve。
 * Promise resovled 不代表编辑器已经完整加载完毕，只代表 SDK 已经准备好了。
 * 同时 Promise 一直 pending 也不代表编辑器加载失败，只代表无法通过 SDK 和编辑器交互。
 * 比如受浏览器限制无法发出 postMessage() 时，Promise 将会一直 pending。
 */
export async function connect(options: ConnectOptions): Promise<OfficeSDK> {
  return await connectInternal(options)
}

/**
 * 初始化 V2 统一编辑 / 预览入口。
 *
 * V2 在运行时强制校验 type 和 mode，底层 iframe 初始化流程与旧 connect 共用。
 */
export async function connectV2(options: ConnectV2Options): Promise<OfficeSDK> {
  validateConnectV2Options(options)

  return await connectInternal(options, 'v2')
}
