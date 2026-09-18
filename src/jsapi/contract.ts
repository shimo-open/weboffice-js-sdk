export type JSAPITransport =
  | 'request'
  | 'subscription'
  | 'handle'
  | 'delegation'

export type JSAPIAdapter =
  | 'none'
  | 'delta-snapshot'
  | 'range-value'
  | 'permission'
  | 'handle'

export interface JSAPIContract {
  id: string
  productPath: string
  facadePath: string
  transport: JSAPITransport
  adapter: JSAPIAdapter
  input: string
  productReturn: string
  facadeReturn: string
  prerequisites: readonly string[]
  sideEffects: readonly string[]
  verification: readonly string[]
}

export interface JSAPIContractCatalog {
  version: number
  suite: string
  sourceType: string
  apis: readonly JSAPIContract[]
}

export function validateJSAPIContractCatalog(
  catalog: JSAPIContractCatalog
): string[] {
  const errors: string[] = []
  const ids = new Set<string>()
  const productPaths = new Set<string>()
  const facadePaths = new Set<string>()

  for (const api of catalog.apis) {
    if (ids.has(api.id)) errors.push(`duplicate id: ${api.id}`)
    ids.add(api.id)
    if (productPaths.has(api.productPath)) {
      errors.push(`duplicate product path: ${api.productPath}`)
    }
    productPaths.add(api.productPath)
    if (facadePaths.has(api.facadePath)) {
      errors.push(`duplicate facade path: ${api.facadePath}`)
    }
    facadePaths.add(api.facadePath)
    const validRoot = ['ActiveOutline.'].some((prefix) =>
      api.facadePath.startsWith(prefix)
    )
    if (!validRoot) {
      errors.push(`facade path must start with a supported root: ${api.id}`)
    }
    if (
      api.transport === 'subscription' &&
      !api.verification.includes('unsubscribe')
    ) {
      errors.push(`subscription must verify unsubscribe: ${api.id}`)
    }
    if (api.sideEffects.length > 0 && api.verification.length === 0) {
      errors.push(`side-effect API must define verification: ${api.id}`)
    }
  }

  return errors
}
