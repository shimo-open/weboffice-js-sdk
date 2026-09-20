import rawCatalog from '../../api-contracts/public-methods.json'

export type PublicMethodTransport =
  | 'product-jsapi'
  | 'subscription'
  | 'delegation-or-product-jsapi'

export interface PublicMethodContract {
  id: string
  publicPath: string
  productPath: string
  facadePath: string
  transport: PublicMethodTransport
}

export interface PublicMethodCatalog {
  version: number
  suite: string
  implementation: string
  methods: readonly PublicMethodContract[]
}

export const PUBLIC_METHODS_CONTRACT = rawCatalog as PublicMethodCatalog
