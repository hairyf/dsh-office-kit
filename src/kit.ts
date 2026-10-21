import { existsSync } from 'node:fs'
import { createRequire } from 'node:module'
import { join, sep } from 'node:path'
import process from 'node:process'
import { fileURLToPath } from 'node:url'

/** Package root of the installed plugin (dist/ during a build, src/ in a checkout). */
export const PACKAGE_ROOT = fileURLToPath(new URL('..', import.meta.url))

const KIT_CLI = '@deepseek-ai/libreoffice-kit/cli'

/**
 * Electron reads `app.asar` as a directory but nothing else can, and the helper
 * runs the recorded path from Python. The packaged app ships the kit unpacked.
 */
function outsideAsar(path: string): string {
  const packed = `${sep}app.asar${sep}`
  if (!path.includes(packed))
    return path
  const unpacked = path.replace(packed, `${sep}app.asar.unpacked${sep}`)
  return existsSync(unpacked) ? unpacked : path
}

/**
 * Locate the LibreOffice kit CLI that DeepSeek Harness ships.
 *
 * DSH installs third-party plugins into a profile (`~/.dsh/profiles/<name>/node_modules`)
 * while the kit lives in the harness dependency tree, so resolving relative to this
 * package cannot find it. Resolving from the harness entry point (`process.argv[1]`,
 * e.g. `<dsh>/node_modules/@deepseek-ai/dsh/lib/bin.js`) walks the tree the harness
 * itself loaded from, and resolving from this package covers a checkout or a
 * plugin-local install. `DSH_LIBREOFFICE_KIT_CLI` overrides both.
 */
export function resolveKitCli(entry: string | undefined = process.argv[1]): string | null {
  const override = process.env.DSH_LIBREOFFICE_KIT_CLI
  if (override != null && existsSync(override))
    return override

  for (const anchor of [entry, join(PACKAGE_ROOT, 'package.json')]) {
    if (anchor == null)
      continue
    try {
      const cli = createRequire(anchor).resolve(KIT_CLI)
      if (existsSync(cli))
        return outsideAsar(cli)
    }
    catch {
      // Not installed at this anchor; try the next one.
    }
  }
  return null
}
