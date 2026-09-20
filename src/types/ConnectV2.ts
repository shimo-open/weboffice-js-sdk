/**
 * V2 统一入口支持的前端套件类型。
 *
 * 当前仅用于 connectV2 参数校验和 iframe 侧资源选择预留，
 * 不参与后端文件创建、导入或编辑准备判断。
 */
export const CONNECT_V2_FILE_TYPES = [
  'file',
  'document',
  'documentPro',
  'writer',
  'spreadsheet',
  'presentation',
  'table',
  'form'
] as const

export type ConnectV2FileType = (typeof CONNECT_V2_FILE_TYPES)[number]

/**
 * Preload 路由版本标识。只有完整的 `v2` 才启用 V2 路径，缺失或其他值按 V1 处理。
 */
export type ContentVersion = 'v1' | 'v2'

export function isConnectV2FileType(
  value: unknown
): value is ConnectV2FileType {
  return (
    typeof value === 'string' &&
    (CONNECT_V2_FILE_TYPES as readonly string[]).includes(value)
  )
}

export function isConnectV2Mode(value: unknown): value is 'edit' | 'preview' {
  return value === 'edit' || value === 'preview'
}

/**
 * 校验 connectV2 的两个新增必填参数。
 */
export function validateConnectV2Options(options: {
  type?: unknown
  mode?: unknown
}): void {
  if (!isConnectV2FileType(options?.type)) {
    throw new Error(`invalid connectV2 "type": ${String(options?.type)}`)
  }

  if (!isConnectV2Mode(options?.mode)) {
    throw new Error(`invalid connectV2 "mode": ${String(options?.mode)}`)
  }
}
