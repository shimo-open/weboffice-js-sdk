const fs = require('fs')
const path = require('path')

const distDir = path.join(__dirname, '..', 'dist')
const declarationDir = path.join(distDir, 'src')

if (!fs.existsSync(declarationDir)) {
  throw new Error(`declaration directory not found: ${declarationDir}`)
}

for (const entry of fs.readdirSync(declarationDir)) {
  const source = path.join(declarationDir, entry)
  const target = path.join(distDir, entry)
  if (fs.existsSync(target)) {
    throw new Error(`declaration target already exists: ${target}`)
  }
  fs.renameSync(source, target)
}

fs.rmdirSync(declarationDir)
