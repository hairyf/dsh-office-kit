import { mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import process from 'node:process'
import { afterEach, describe, expect, it } from 'vitest'
import { resolveKitCli } from '../src/kit'
import { saveKitPath } from '../src/python'

const temporary: string[] = []
const originalOverride = process.env.DSH_LIBREOFFICE_KIT_CLI

function sandbox(): string {
  const root = realpathSync(mkdtempSync(join(tmpdir(), 'dsh-office-kit-')))
  temporary.push(root)
  return root
}

/** Lay out `node_modules/@deepseek-ai/libreoffice-kit` plus a harness entry inside it. */
function fakeHarness(root: string): string {
  const kit = join(root, 'node_modules', '@deepseek-ai', 'libreoffice-kit')
  mkdirSync(join(kit, 'lib'), { recursive: true })
  writeFileSync(join(kit, 'package.json'), JSON.stringify({
    name: '@deepseek-ai/libreoffice-kit',
    version: '9.9.9',
    exports: { './cli': './lib/cli.js' },
  }))
  writeFileSync(join(kit, 'lib', 'cli.js'), '')
  const entry = join(root, 'node_modules', '@deepseek-ai', 'dsh', 'lib', 'bin.js')
  mkdirSync(join(entry, '..'), { recursive: true })
  writeFileSync(entry, '')
  return entry
}

afterEach(() => {
  for (const root of temporary.splice(0))
    rmSync(root, { recursive: true, force: true })
  if (originalOverride == null)
    delete process.env.DSH_LIBREOFFICE_KIT_CLI
  else
    process.env.DSH_LIBREOFFICE_KIT_CLI = originalOverride
})

describe('resolveKitCli', () => {
  it('resolves the kit from the tree the harness entry point runs in', () => {
    delete process.env.DSH_LIBREOFFICE_KIT_CLI
    const root = sandbox()
    const entry = fakeHarness(root)
    expect(resolveKitCli(entry)).toBe(join(root, 'node_modules', '@deepseek-ai', 'libreoffice-kit', 'lib', 'cli.js'))
  })

  it('honours DSH_LIBREOFFICE_KIT_CLI before any lookup', () => {
    const root = sandbox()
    const cli = join(root, 'custom-cli.js')
    writeFileSync(cli, '')
    process.env.DSH_LIBREOFFICE_KIT_CLI = cli
    expect(resolveKitCli(join(root, 'missing', 'bin.js'))).toBe(cli)
  })
})

describe('saveKitPath', () => {
  it('records the kit beside the interpreter postinstall saved', () => {
    const root = sandbox()
    const config = join(root, 'config.json')
    writeFileSync(config, JSON.stringify({ python: 'C:/venv/python.exe', version: '0.4.1' }))
    saveKitPath(config, 'C:/node/node.exe', 'C:/kit/lib/cli.js')
    // A legacy config on the machine may contribute extra keys, so match, not equal.
    expect(JSON.parse(readFileSync(config, 'utf8'))).toMatchObject({
      python: 'C:/venv/python.exe',
      version: '0.4.1',
      libreofficeKit: { node: 'C:/node/node.exe', cli: 'C:/kit/lib/cli.js' },
    })
    const first = readFileSync(config, 'utf8')
    saveKitPath(config, 'C:/node/node.exe', 'C:/kit/lib/cli.js')
    expect(readFileSync(config, 'utf8')).toBe(first)
  })
})
