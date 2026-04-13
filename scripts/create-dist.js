#!/usr/bin/env node
/**
 * create-dist.js
 * Builds ExhiBytes and creates a distributable ZIP:
 *
 *   ExhiBytes-v{version}-Setup.zip
 *   ├── app\                              (ExhiBytes program files)
 *   ├── LibreOffice_26.2.2_Win_x86-64.msi
 *   ├── Install.bat
 *   └── README.txt
 *
 * Usage:  npm run dist
 */

const { execSync, spawnSync } = require('child_process')
const fs   = require('fs')
const path = require('path')
const os   = require('os')

const ROOT        = path.join(__dirname, '..')
const DIST_DIR    = path.join(ROOT, 'dist')
const UNPACKED    = path.join(DIST_DIR, 'win-unpacked')
const RESOURCES   = path.join(ROOT, 'resources')
const LO_MSI      = path.join(RESOURCES, 'LibreOffice_26.2.2_Win_x86-64.msi')
const INSTALL_BAT = path.join(ROOT, 'Install.bat')
const README      = path.join(ROOT, 'README.txt')
const STAGING     = path.join(DIST_DIR, '_staging')

// ── Helpers ──────────────────────────────────────────────────────────────────

function run(cmd) {
  console.log(`\n> ${cmd}`)
  execSync(cmd, { stdio: 'inherit', cwd: ROOT })
}

function readVersion() {
  return JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version || '1.0.0'
}

function rmrf(dir) {
  if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true })
}

// ── Fix winCodeSign cache ─────────────────────────────────────────────────────
// electron-builder downloads winCodeSign (contains NSIS + signtool).
// The .7z has macOS symlinks that 7-Zip cannot create on Windows without
// SeCreateSymbolicLinkPrivilege (requires admin or Developer Mode).
// Workaround: take the partially-extracted temp dir (all Windows files are
// present), stub the 2 missing macOS dylib symlinks, and move to final location.

function fixWinCodeSignCache() {
  const cacheEnv = process.env.ELECTRON_BUILDER_CACHE
  const cacheBase = cacheEnv
    ? path.join(cacheEnv, 'winCodeSign')
    : path.join(os.homedir(), 'AppData', 'Local', 'electron-builder', 'Cache', 'winCodeSign')

  const finalCache = path.join(cacheBase, 'winCodeSign-2.6.0')
  if (fs.existsSync(finalCache)) {
    console.log('  winCodeSign cache already valid — skipping fix')
    return true
  }

  if (!fs.existsSync(cacheBase)) return false

  // Find a partial temp extraction (numeric directory names)
  const candidates = fs.readdirSync(cacheBase)
    .filter(e => /^\d+$/.test(e) && fs.statSync(path.join(cacheBase, e)).isDirectory())
    .map(e => ({ name: e, p: path.join(cacheBase, e), mtime: fs.statSync(path.join(cacheBase, e)).mtimeMs }))
    .sort((a, b) => b.mtime - a.mtime)

  if (candidates.length === 0) return false

  const tempDir = candidates[0].p
  // Verify it has the expected Windows content
  if (!fs.existsSync(path.join(tempDir, 'windows-10'))) return false

  // Stub the 2 missing macOS dylib symlinks
  for (const stub of [
    path.join(tempDir, 'darwin', '10.12', 'lib', 'libcrypto.dylib'),
    path.join(tempDir, 'darwin', '10.12', 'lib', 'libssl.dylib')
  ]) {
    if (!fs.existsSync(stub)) {
      fs.mkdirSync(path.dirname(stub), { recursive: true })
      fs.writeFileSync(stub, '') // empty stub — not needed on Windows
    }
  }

  // Move to the final cache location
  fs.renameSync(tempDir, finalCache)
  console.log(`  ✓  winCodeSign cache fixed → ${finalCache}`)
  return true
}

// ── Main ─────────────────────────────────────────────────────────────────────

async function main() {
  console.log('=== ExhiBytes Distribution Builder ===\n')

  // 1. Build renderer + main
  console.log('Step 1/4  Building app…')
  run('npm run build')

  // 2. Package with electron-builder (dir target — no NSIS installer, just the app folder)
  console.log('\nStep 2/4  Packaging with electron-builder…')

  let packed = false
  for (let attempt = 1; attempt <= 2; attempt++) {
    try {
      run('npx electron-builder --win dir --publish never')
      packed = true
      break
    } catch (e) {
      if (attempt === 1) {
        console.log('\n  electron-builder failed (likely winCodeSign symlink error on Windows).')
        console.log('  Applying cache fix and retrying…')
        const fixed = fixWinCodeSignCache()
        if (!fixed) {
          console.error('\n  Could not fix winCodeSign cache automatically.')
          console.error('  Please enable Windows Developer Mode (Settings → System → Developer Mode)')
          console.error('  or run this script from an Administrator terminal, then retry.')
          process.exit(1)
        }
      } else {
        throw e
      }
    }
  }

  if (!packed || !fs.existsSync(UNPACKED)) {
    console.error(`ERROR: Expected output not found: ${UNPACKED}`)
    process.exit(1)
  }

  // 2b. Explicitly embed icon into exe using rcedit.
  // electron-builder's --win dir mode does not always invoke rcedit on this
  // platform, so the icon in ExhiBytes.exe can remain as the default Electron
  // icon. We run rcedit ourselves here to guarantee the correct icon is set.
  console.log('\n  Embedding icon into ExhiBytes.exe…')
  const exePath  = path.join(UNPACKED, 'ExhiBytes.exe')
  const icoPath  = path.join(ROOT, 'resources', 'icon.ico')
  const cacheBase = path.join(os.homedir(), 'AppData', 'Local', 'electron-builder', 'Cache', 'winCodeSign')
  let rcedit = null
  if (fs.existsSync(cacheBase)) {
    // Find the most recently modified winCodeSign directory that has rcedit-x64.exe
    const dirs = fs.readdirSync(cacheBase)
      .map(d => path.join(cacheBase, d))
      .filter(d => {
        try { return fs.statSync(d).isDirectory() } catch { return false }
      })
    for (const d of dirs.reverse()) {
      const candidate = path.join(d, 'rcedit-x64.exe')
      if (fs.existsSync(candidate)) { rcedit = candidate; break }
    }
  }
  if (rcedit && fs.existsSync(exePath) && fs.existsSync(icoPath)) {
    const r = spawnSync(rcedit, ['--set-icon', icoPath, exePath], { stdio: 'pipe' })
    if (r.status === 0) {
      console.log('  ✓  Icon embedded via rcedit')
    } else {
      console.warn('  ⚠  rcedit failed:', (r.stderr || r.stdout || '').toString().trim())
    }
  } else {
    console.warn('  ⚠  rcedit or target files not found — icon embedding skipped')
  }

  // 3. Assemble staging folder
  console.log('\nStep 3/4  Assembling distribution folder…')
  rmrf(STAGING)
  fs.mkdirSync(STAGING, { recursive: true })

  // Copy app directory
  const appDest = path.join(STAGING, 'app')
  fs.mkdirSync(appDest, { recursive: true })
  execSync(`xcopy /E /Y /I /Q "${UNPACKED}" "${appDest}\\"`, { stdio: 'inherit' })
  console.log('  ✓  app\\')

  // Copy LibreOffice MSI
  if (fs.existsSync(LO_MSI)) {
    fs.copyFileSync(LO_MSI, path.join(STAGING, path.basename(LO_MSI)))
    console.log(`  ✓  ${path.basename(LO_MSI)}  (${(fs.statSync(LO_MSI).size / 1024 / 1024).toFixed(0)} MB)`)
  } else {
    console.warn(`  ⚠  LibreOffice MSI not found at resources/${path.basename(LO_MSI)} — omitting`)
  }

  // Copy Install.bat
  if (fs.existsSync(INSTALL_BAT)) {
    fs.copyFileSync(INSTALL_BAT, path.join(STAGING, 'Install.bat'))
    console.log('  ✓  Install.bat')
  } else {
    console.warn('  ⚠  Install.bat not found — omitting')
  }

  // Copy README
  if (fs.existsSync(README)) {
    fs.copyFileSync(README, path.join(STAGING, 'README.txt'))
    console.log('  ✓  README.txt')
  }

  // 4. Create ZIP using PowerShell Compress-Archive
  console.log('\nStep 4/4  Creating ZIP…')
  const version = readVersion()
  const zipName = `ExhiBytes-v${version}-Setup.zip`
  const zipPath = path.join(DIST_DIR, zipName)

  if (fs.existsSync(zipPath)) fs.unlinkSync(zipPath)

  const psCmd = [
    `Compress-Archive -Path '${STAGING.replace(/'/g, "''")}\\*'`,
    ` -DestinationPath '${zipPath.replace(/'/g, "''")}'`,
    ` -CompressionLevel Optimal`
  ].join('')

  const result = spawnSync('powershell', ['-NoProfile', '-Command', psCmd], {
    stdio: 'inherit', cwd: ROOT
  })

  if (result.status !== 0) {
    console.error('\nERROR: Compress-Archive failed.')
    process.exit(1)
  }

  // Cleanup staging
  rmrf(STAGING)

  const sizeMB = (fs.statSync(zipPath).size / 1024 / 1024).toFixed(1)
  console.log(`\n✓  Done!  Distribution package:`)
  console.log(`   ${zipPath}`)
  console.log(`   Size: ${sizeMB} MB`)
  console.log('\n   Recipients: extract the ZIP, then right-click Install.bat → Run as administrator.')
}

main().catch(err => { console.error('\n' + err.message); process.exit(1) })
