const fs = require('fs')
const path = require('path')

const repoRoot = path.resolve(__dirname, '..')
const catalogPath = path.join(repoRoot, 'api-contracts', 'light-doc-p0.json')
const catalog = JSON.parse(fs.readFileSync(catalogPath, 'utf8'))
const errors = []
const requireShowcase = process.argv.includes('--require-showcase')

const ids = new Set()
const productPaths = new Set()
const facadePaths = new Set()
const allowedTransports = new Set([
  'request',
  'subscription',
  'handle',
  'delegation'
])
const allowedAdapters = new Set([
  'none',
  'delta-snapshot',
  'range-value',
  'permission',
  'handle'
])

for (const api of catalog.apis) {
  if (ids.has(api.id)) errors.push(`duplicate id: ${api.id}`)
  ids.add(api.id)
  if (productPaths.has(api.productPath))
    errors.push(`duplicate product path: ${api.productPath}`)
  productPaths.add(api.productPath)
  if (facadePaths.has(api.facadePath))
    errors.push(`duplicate facade path: ${api.facadePath}`)
  facadePaths.add(api.facadePath)
  const validRoot = api.facadePath.startsWith('ActiveDocument.')
  if (!validRoot)
    errors.push(`facade path must start with a supported root: ${api.id}`)
  if (!allowedTransports.has(api.transport))
    errors.push(`invalid transport: ${api.id}`)
  if (!allowedAdapters.has(api.adapter))
    errors.push(`invalid adapter: ${api.id}`)
  if (!Array.isArray(api.prerequisites))
    errors.push(`missing prerequisites: ${api.id}`)
  if (!Array.isArray(api.sideEffects))
    errors.push(`missing sideEffects: ${api.id}`)
  if (!Array.isArray(api.verification) || api.verification.length === 0)
    errors.push(`missing verification: ${api.id}`)
  if (
    api.transport === 'subscription' &&
    !api.verification.includes('unsubscribe')
  ) {
    errors.push(`subscription must verify unsubscribe: ${api.id}`)
  }
}

const sourceFiles = [
  path.join(repoRoot, 'src', 'OfficeSDK.facade.ts'),
  path.join(repoRoot, 'src', 'OfficeSDK.facade.types.ts'),
  path.join(repoRoot, 'src', 'editorFacade.contract.ts')
]
const sourceText = sourceFiles
  .map((file) => fs.readFileSync(file, 'utf8'))
  .join('\n')
for (const api of catalog.apis) {
  const methodName = api.facadePath.split('.').pop()
  if (!sourceText.includes(methodName))
    errors.push(`facade method is not declared in SDK source: ${api.id}`)
}

const iframeRoot = process.env.JSAPI_IFRAME_ROOT
if (iframeRoot) {
  const iframeFile = path.join(
    iframeRoot,
    'packages',
    'editor-sdk-runtime',
    'src',
    'components',
    'NewDoc',
    'product-jsapi.ts'
  )
  if (fs.existsSync(iframeFile)) {
    const iframeText = fs.readFileSync(iframeFile, 'utf8')
    for (const api of catalog.apis) {
      if (!iframeText.includes(`'${api.productPath}'`))
        errors.push(`receiver path is not registered: ${api.id}`)
    }
  } else {
    console.warn(
      '[contract-check] JSAPI_IFRAME_ROOT is configured but receiver file was not found; skipped receiver scan'
    )
  }
} else {
  console.warn(
    '[contract-check] JSAPI_IFRAME_ROOT is not configured; skipped receiver scan'
  )
}

const showcaseRoot = process.env.JSAPI_SHOWCASE_ROOT
if (showcaseRoot) {
  const showcaseFile = path.join(
    showcaseRoot,
    'ui',
    'src',
    'pages',
    'LightDocJsApiPage',
    'LightDocJsApiPage.registry.ts'
  )
  if (fs.existsSync(showcaseFile)) {
    const showcaseText = fs.readFileSync(showcaseFile, 'utf8')
    for (const api of catalog.apis) {
      if (!showcaseText.includes(`key: '${api.id}'`)) {
        const message = `Showcase entry is not registered: ${api.id}`
        if (requireShowcase) errors.push(message)
        else
          console.warn(
            `[contract-check] ${message} (business-layer check skipped)`
          )
      }
    }
  } else {
    const message =
      'JSAPI_SHOWCASE_ROOT is configured but Showcase registry was not found'
    if (requireShowcase) errors.push(message)
    else
      console.warn(`[contract-check] ${message} (business-layer check skipped)`)
  }
} else if (requireShowcase) {
  errors.push('JSAPI_SHOWCASE_ROOT is required for showcase-contract-check')
} else {
  console.warn(
    '[contract-check] JSAPI_SHOWCASE_ROOT is not configured; skipped Showcase scan'
  )
}

if (errors.length > 0) {
  console.error('[contract-check] failed')
  for (const error of errors) console.error(`- ${error}`)
  process.exitCode = 1
} else {
  console.log(
    `[contract-check] passed: ${catalog.apis.length} APIs${
      requireShowcase ? ' (including Showcase)' : ''
    }`
  )
}
