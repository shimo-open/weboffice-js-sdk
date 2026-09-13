import type { ConnectV2FileType, ContentVersion } from './types/ConnectV2'

export interface PreloadInitPayload {
  token?: string
  signature?: string
  fileGuid?: string
  mode?: 'edit' | 'preview'
  type?: ConnectV2FileType
  contentVersion?: ContentVersion
}

export interface CreatePreloadInitPayloadOptions extends PreloadInitPayload {
  contentVersion?: ContentVersion
}

export function createPreloadInitPayload(
  options: CreatePreloadInitPayloadOptions
): PreloadInitPayload {
  const payload: PreloadInitPayload = {
    token: options.token,
    signature: options.signature,
    fileGuid: options.fileGuid,
    mode: options.mode,
    type: options.type
  }

  if (options.contentVersion === 'v2') {
    payload.contentVersion = 'v2'
  }

  return payload
}
