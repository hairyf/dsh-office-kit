import type { OfficeFeatureId } from '../src/features'
import { Buffer } from 'node:buffer'
import { execFileSync } from 'node:child_process'
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import process from 'node:process'
import { afterAll, describe, expect, it } from 'vitest'
import { featureByCommand, featureById, OFFICE_FEATURES } from '../src/features'
import { officePaths, venvPython } from '../src/python'

const PACKAGE_ROOT = new URL('..', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/u, '$1')

function configuredPython(): string | undefined {
  const candidates: (string | undefined)[] = [process.env.DSH_OFFICE_PYTHON]
  // Read the same paths the host half resolves. Hardcoding the Windows-shaped
  // config location here made every integration test skip on Linux.
  const paths = officePaths()
  for (const config of [paths.configFile, ...paths.legacyConfigFiles]) {
    if (!existsSync(config))
      continue
    try {
      const parsed = JSON.parse(readFileSync(config, 'utf8')) as { python?: unknown }
      if (typeof parsed.python === 'string')
        candidates.push(parsed.python)
    }
    catch {
      // Ignore a malformed config; the test falls back to the checked-in venv.
    }
  }
  candidates.push(venvPython(paths.venv))
  candidates.push('python3', 'python')
  return candidates.find(candidate => candidate != null && (candidate.includes('/') || candidate.includes('\\') ? existsSync(candidate) : true))
}

interface CliStatus {
  readonly ok: boolean
  readonly plugin_version: string
  readonly packages: readonly { package: string, installed: boolean }[]
  readonly renderers: {
    readonly kit: {
      readonly available: boolean
      readonly backend: string | null
      readonly version: string | null
      readonly error?: string
      readonly remedy?: string
    }
    readonly can_render: { docx: boolean, pptx: boolean, xlsx: boolean, pdf: boolean }
    readonly can_recalculate: { xlsx: boolean }
  }
}

function runJson(command: string, args: string[]): Record<string, any> {
  const stdout = execFileSync(command, args, {
    encoding: 'utf8',
    env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
    maxBuffer: 8 * 1024 * 1024,
  })
  return JSON.parse(stdout)
}

function runCli(python: string, args: string[]): Record<string, any> {
  return runJson(python, [officePaths().cli, ...args])
}

function runScript(python: string, pack: string, script: string, args: string[]): Record<string, any> {
  return runJson(python, [join(PACKAGE_ROOT, 'assets', 'skills', pack, script), ...args])
}

/** Run a script that may use exit code 1 as a report result (audits). */
function runScriptAudit(python: string, pack: string, script: string, args: string[]): { status: number, payload: Record<string, any> } {
  try {
    const stdout = execFileSync(python, [join(PACKAGE_ROOT, 'assets', 'skills', pack, script), ...args], {
      encoding: 'utf8',
      env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
      maxBuffer: 8 * 1024 * 1024,
    })
    return { status: 0, payload: JSON.parse(stdout) }
  }
  catch (error) {
    const failure = error as { status?: number, stdout?: string }
    return { status: failure.status ?? 1, payload: JSON.parse(String(failure.stdout ?? '{}')) }
  }
}

/** Run a pack script the way a user shell would: no UTF-8 env overrides, POSIX locale. */
function runScriptRaw(python: string, pack: string, script: string, args: string[]): Buffer {
  const env: NodeJS.ProcessEnv = { ...process.env, LC_ALL: 'C', LANG: 'C' }
  delete env.PYTHONIOENCODING
  delete env.PYTHONUTF8
  return execFileSync(python, [join(PACKAGE_ROOT, 'assets', 'skills', pack, script), ...args], { env, maxBuffer: 8 * 1024 * 1024 })
}

const python = configuredPython()
let status: CliStatus | undefined
if (python != null) {
  try {
    status = runCli(python, ['status']) as unknown as CliStatus
  }
  catch {
    status = undefined
  }
}
const ready = status != null && status.packages.every(item => item.installed)

// CI sets this so a platform without a working kit fails instead of quietly skipping.
const requireRenderer = process.env.DSH_OFFICE_REQUIRE_RENDERER === '1'
const canRenderDocx = requireRenderer || status?.renderers.can_render.docx === true
const canRecalculate = requireRenderer || status?.renderers.can_recalculate.xlsx === true

describe('feature registry', () => {
  it('exposes four stable ids and slash commands', () => {
    expect(OFFICE_FEATURES.map(feature => feature.id)).toEqual(['documents', 'pdf', 'presentations', 'spreadsheets'])
    expect(OFFICE_FEATURES.map(feature => feature.command)).toEqual(['documents', 'pdf', 'presentations', 'spreadsheets'])
    for (const feature of OFFICE_FEATURES) {
      expect(featureById(feature.id)).toBe(feature)
      expect(featureByCommand(feature.command)).toBe(feature)
      expect(feature.skill).toMatch(/^office-/u)
      expect(feature.label.length).toBeGreaterThan(0)
    }
  })

  it('returns undefined for unknown ids and commands', () => {
    expect(featureById('slides')).toBeUndefined()
    expect(featureByCommand('docs')).toBeUndefined()
  })
})

const CLI_KIND: Record<OfficeFeatureId, 'docx' | 'pdf' | 'pptx' | 'xlsx'> = {
  documents: 'docx',
  pdf: 'pdf',
  presentations: 'pptx',
  spreadsheets: 'xlsx',
}

describe.runIf(ready)('office CLI integration', () => {
  const directory = mkdtempSync(join(tmpdir(), 'dsh office 中文 '))
  const image = join(directory, 'icon.png')
  const specs: Record<'docx' | 'pdf' | 'pptx' | 'xlsx', object> = {
    docx: {
      title: '季度报告 Q3',
      author: 'DSH Office Suite',
      blocks: [
        { type: 'heading', level: 1, text: '执行摘要' },
        { type: 'paragraph', text: '本季度收入同比增长 18.4%。' },
        { type: 'bullets', items: ['续约率 92%', '新增客户 42 家'] },
        { type: 'table', headers: ['指标', 'Q3'], rows: [['收入', '¥12,480,000'], ['毛利率', '71.2%']] },
        { type: 'equation', latex: '\\frac{a^2+b^2}{c^2}' },
      ],
    },
    pdf: {
      title: '产品说明',
      blocks: [
        { type: 'heading', level: 1, text: '产品说明' },
        { type: 'paragraph', text: '中文、English、数字 123 与标点。' },
        { type: 'table', headers: ['项目', '值'], rows: [['版本', '1.0']] },
      ],
    },
    pptx: {
      title: '年度汇报',
      slides: [
        { layout: 'title', title: '年度汇报', subtitle: '内部审阅' },
        { layout: 'blank', title: '数据', table: { headers: ['季度', '收入'], rows: [['Q1', 100], ['Q2', 128]] }, chart: { type: 'column', left: 3.6, top: 3.5, width: 8.8, height: 3.0, categories: ['Q1', 'Q2'], series: [{ name: '收入', values: [100, 128] }] } },
      ],
    },
    xlsx: {
      sheets: [{
        name: '汇总',
        rows: [['季度', '收入', '成本', '毛利'], ['Q1', 100, 60, '=B2-C2'], ['Q2', 128, 70, '=B3-C3'], [null, '=SUM(B2:B3)']],
        freeze: 'A2',
        styles: { 'A1:D1': { bold: true, fill: '#E8EEF7' } },
        charts: [{ type: 'column', anchor: 'F2', title: '收入趋势', data: { min_col: 2, min_row: 1, max_col: 2, max_row: 3 } }],
      }],
    },
  }

  writeFileSync(image, Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64'))

  function create(feature: OfficeFeatureId, name: string): string {
    const kind = CLI_KIND[feature]
    const spec = join(directory, `${kind}.json`)
    const out = join(directory, `${name}.${kind}`)
    writeFileSync(spec, JSON.stringify(specs[kind]))
    const result = runCli(python!, [kind, 'create', '--spec', spec, '--out', out])
    expect(result.ok).toBe(true)
    expect(existsSync(out)).toBe(true)
    return out
  }

  function read(kind: string, file: string, dataOnly = false): Record<string, any> {
    return runCli(python!, [kind, 'read', '--input', file, ...(dataOnly ? ['--data-only'] : [])])
  }

  afterAll(() => rmSync(directory, { recursive: true, force: true }))

  it('reports the version from package.json', () => {
    const manifest = JSON.parse(readFileSync(join(PACKAGE_ROOT, 'package.json'), 'utf8')) as { version: string }
    expect(status?.plugin_version).toBe(manifest.version)
  })

  it('rejects the edit verb that now lives in the pack scripts', () => {
    expect(() => runCli(python!, ['docx', 'edit', '--help'])).toThrow()
  })

  it('creates and reads DOCX content including CJK, tables, and equations', () => {
    const file = create('documents', '季度 报告')
    const result = read('docx', file)
    expect(result.headings).toContain('执行摘要')
    expect(result.tables[0].rows[1][0]).toBe('收入')
    expect(result.equations).toBe(1)
  })

  it('edits DOCX text and appends a heading through the pack scripts', () => {
    const file = create('documents', '季度 报告-编辑前')
    const replaced = join(directory, '季度 报告-替换.docx')
    const edited = join(directory, '季度 报告-编辑后.docx')
    expect(runScript(python!, 'office-documents', 'scripts/replace_text.py', ['--input', file, '--find', '42 家', '--replace', '51 家', '--out', replaced]).replacements).toBe(1)
    expect(runScript(python!, 'office-documents', 'scripts/append_blocks.py', ['--input', replaced, '--blocks', JSON.stringify([{ type: 'heading', level: 2, text: '附录' }]), '--out', edited]).appended).toBe(1)
    const result = read('docx', edited)
    expect(result.paragraphs.some((item: { text: string }) => item.text.includes('51 家'))).toBe(true)
    expect(result.headings).toContain('附录')
    expect(result.equations).toBe(1)
  })

  it('writes DOCX without library metadata and marks the table header row', () => {
    const file = create('documents', '元数据 检查')
    const report = runJson(python!, ['-c', 'import json, sys; from docx import Document; d = Document(sys.argv[1]); p = d.core_properties; print(json.dumps({"comments": p.comments, "last_modified_by": p.last_modified_by, "header": "w:tblHeader" in d.tables[0].rows[0]._tr.trPr.xml}))', file])
    expect(report.comments).toBe('')
    expect(report.last_modified_by).toBe('DSH Office Suite')
    expect(report.header).toBe(true)
  })

  it('creates and reads a Chinese PDF', () => {
    const file = create('pdf', '产品 说明')
    const result = read('pdf', file)
    expect(result.page_count).toBeGreaterThan(0)
    expect(result.pages.map((page: { text: string }) => page.text).join('')).toContain('产品说明')
  })

  it('never emits NUL text when the PDF spec contains CJK', () => {
    // A Latin-only fallback font (DejaVuSans) registers cleanly and writes each CJK
    // character as NUL, so a plain `ok: true` is not evidence the text survived.
    const file = create('pdf', '编码 检查')
    const text = read('pdf', file).pages.map((page: { text: string }) => page.text).join('')
    expect(text).toContain('产品说明')
    expect(text).not.toContain('\u0000')
  })

  it('renders a PDF title once when the spec also carries a title block', () => {
    const spec = join(directory, 'pdf-title.json')
    const out = join(directory, '巡检 报告.pdf')
    writeFileSync(spec, JSON.stringify({ title: '巡检报告', blocks: [{ type: 'title', text: '巡检报告' }, { type: 'paragraph', text: '正文' }] }))
    expect(runCli(python!, ['pdf', 'create', '--spec', spec, '--out', out]).ok).toBe(true)
    const text = read('pdf', out).pages.map((page: { text: string }) => page.text).join('')
    expect(text.split('巡检报告')).toHaveLength(2)
  })

  it('writes PDF metadata through the pack script and leaves the original untouched', () => {
    const file = create('pdf', '产品 说明-编辑前')
    const out = join(directory, '产品 说明-编辑后.pdf')
    const report = runScript(python!, 'office-pdf', 'scripts/metadata.py', ['--input', file, '--set', 'title=产品说明（已编辑）', '--out', out])
    expect(report.applied.title).toBe('产品说明（已编辑）')
    expect(read('pdf', out).metadata['/Title']).toBe('产品说明（已编辑）')
    expect(read('pdf', file).metadata['/Title']).not.toBe('产品说明（已编辑）')
  })

  it('merges, drops, and rotates PDF pages through the pack page_ops script', () => {
    const file = create('pdf', '产品 说明-合并')
    const merged = join(directory, '合并.pdf')
    const edited = join(directory, '合并-删除并旋转.pdf')
    expect(runScript(python!, 'office-pdf', 'scripts/page_ops.py', ['--input', file, '--merge', file, '--out', merged]).pages_out).toBe(2)
    expect(runScript(python!, 'office-pdf', 'scripts/page_ops.py', ['--input', merged, '--remove', '0', '--rotate', '90', '--out', edited]).pages_out).toBe(1)
  })

  it('creates, reads, and extends a PPTX deck through the pack script', () => {
    const file = create('presentations', '年度 汇报')
    const result = read('pptx', file)
    expect(result.slide_count).toBe(2)
    expect(result.slides[1].charts[0].categories).toEqual(['Q1', 'Q2'])
    const spec = join(directory, 'pptx-extra-slides.json')
    const out = join(directory, '年度 汇报-编辑后.pptx')
    writeFileSync(spec, JSON.stringify({ slides: [{ layout: 'bullets', title: '下一步', bullets: ['验收', '发布'] }] }))
    expect(runScript(python!, 'office-presentations', 'container_tools/add_slides.py', ['--input', file, '--spec', spec, '--out', out]).added).toBe(1)
    expect(read('pptx', out).slide_count).toBe(3)
  })

  it('creates, reads, and writes cells through the pack script', () => {
    const file = create('spreadsheets', '财务 模型')
    const result = read('xlsx', file)
    expect(result.sheets[0].formulas.length).toBe(3)
    const out = join(directory, '财务 模型-编辑后.xlsx')
    expect(runScript(python!, 'office-spreadsheets', 'container_tools/set_cells.py', ['--input', file, '--sheet', '汇总', '--set', 'E2=已审核', '--out', out]).written).toHaveLength(1)
    expect(read('xlsx', out).sheets[0].rows[1][4]).toBe('已审核')
  })

  it('rasterises a PDF page to a PNG through the shared script', () => {
    const out = join(directory, 'raster', 'page-1.png')
    const result = runScript(python!, 'shared', 'scripts/ensure_raster_image.py', ['--input', join(PACKAGE_ROOT, 'public', 'samples', 'sample.pdf'), '--page', '1', '--out', out]) as { ok: boolean, width: number, height: number }
    expect(result.ok).toBe(true)
    expect(result.width).toBeGreaterThan(500)
    expect(existsSync(out)).toBe(true)
  })

  it('stamps Chinese text by registering a system CJK font', () => {
    const file = create('pdf', '盖章 来源')
    const stamped = join(directory, '盖章 中文.pdf')
    const result = runScript(python!, 'office-pdf', 'scripts/stamp_overlay.py', ['--input', file, '--out', stamped, '--text', '机密 · 第 {page} 页 / 共 {pages} 页', '--position', 'bottom-center'])
    expect(result.ok).toBe(true)
    expect(result.stamped_pages.map((page: { page: number }) => page.page)).toEqual([1])
    expect(result.stamped_pages[0].text).toContain('机密')
    expect(String(result.font_file)).toMatch(/\.tt[cf]$|\.otf$/iu)
    const out = join(directory, '盖章 渲染', 'page-1.png')
    expect(runScript(python!, 'office-pdf', 'scripts/stamp_overlay.py', ['--input', file, '--out', join(directory, '盖章 中文-渲染.pdf'), '--text', '机密', '--render-dir', join(directory, '盖章 渲染')]).rendered_pages.length).toBe(1)
    expect(readFileSync(out).byteLength).toBeGreaterThan(5_000)
  })

  it('paints Chinese form values into the widget appearance and flattens them', () => {
    const form = join(directory, '中文 表单.pdf')
    const maker = join(directory, 'make-form.py')
    writeFileSync(maker, [
      'import sys',
      'from reportlab.pdfgen import canvas',
      'c = canvas.Canvas(sys.argv[1], pagesize=(400, 300))',
      'c.drawString(40, 250, "AcroForm test")',
      'c.acroForm.textfield(name="full_name", x=40, y=200, width=300, height=28)',
      'c.acroForm.textfield(name="company", x=40, y=150, width=300, height=28)',
      'c.acroForm.checkbox(name="agree", x=40, y=110, buttonStyle="check")',
      'c.save()',
      '',
    ].join('\n'))
    execFileSync(python!, [maker, form])

    const values = JSON.stringify({ full_name: '张伟', company: '示例公司', agree: true })
    const filled = join(directory, '中文 表单-已填.pdf')
    const result = runScript(python!, 'office-pdf', 'scripts/forms.py', ['fill', '--input', form, '--out', filled, '--values', values, '--render-dir', join(directory, '渲染-表单')])
    // pypdf alone would paint "?" here; the report names the fields it repainted.
    expect(result.ok).toBe(true)
    expect(result.cjk_appearance_fields).toEqual(['full_name', 'company'])
    expect(result.written.full_name).toBe('张伟')
    expect(readFileSync(join(directory, '渲染-表单', 'page-1.png')).byteLength).toBeGreaterThan(5_000)

    const flat = join(directory, '中文 表单-扁平.pdf')
    const flatten = runScript(python!, 'office-pdf', 'scripts/forms.py', ['fill', '--input', form, '--out', flat, '--values', values, '--flatten'])
    expect(flatten.ok).toBe(true)
    expect(flatten.flatten).toBe(true)
    for (const check of flatten.checks as { check: string, ok: boolean }[]) {
      if (['no widgets remain', 'no form fields remain', '/AcroForm removed'].includes(check.check))
        expect(check.ok, check.check).toBe(true)
    }
  })

  it('writes UTF-8 JSON without UTF-8 env overrides and rasterises SVG', () => {
    const svg = join(directory, '图标.svg')
    const png = join(directory, '图标.png')
    writeFileSync(svg, '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="60"><rect width="120" height="60" fill="#1f3864"/></svg>')
    const raw = runScriptRaw(python!, 'shared', 'scripts/ensure_raster_image.py', ['--input', svg, '--width', '240', '--out', png])
    expect(Buffer.from(raw.toString('utf8'), 'utf8').equals(raw)).toBe(true)
    const result = JSON.parse(raw.toString('utf8'))
    expect(result.ok).toBe(true)
    expect([result.width, result.height]).toEqual([240, 120])
    expect(existsSync(png)).toBe(true)
  })

  it('reports the live-control skill as unusable instead of crashing', () => {
    // Excel/WPS live control is Windows COM territory; every other platform, and a
    // Windows box with nothing open, must fail with one JSON object, not a traceback.
    const script = join(PACKAGE_ROOT, 'assets', 'skills', 'excel-live-control', 'container_tools', 'live_excel.py')
    let exitCode = 0
    let stdout = ''
    try {
      stdout = execFileSync(python!, [script, 'info'], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] })
    }
    catch (error) {
      const failure = error as { status?: number, stdout?: string }
      exitCode = failure.status ?? 0
      stdout = String(failure.stdout ?? '')
    }
    const payload = JSON.parse(stdout)
    expect(payload.ok).toBe(false)
    expect(payload.stage).toBe(process.platform === 'win32' ? 'attach' : 'powershell')
    expect(exitCode).toBe(process.platform === 'win32' ? 1 : 3)
    if (process.platform === 'win32')
      expect(String(payload.remediation)).toMatch(/Excel|WPS/u)
    else
      expect(String(payload.error)).toMatch(/Windows PowerShell 5\.1/u)
  }, 60_000)

  it.runIf(ready)('refreshes a stale interpreter record after the venv moves', async () => {
    const config = join(directory, 'stale-config.json')
    writeFileSync(config, JSON.stringify({
      python: join(directory, 'gone', 'python.exe'),
      libreofficeKit: { node: process.execPath, cli: 'kept' },
    }))
    const { inspectRuntime } = await import('../src/python')
    const resolved = await inspectRuntime(config)
    expect(resolved.state).toBe('ok')
    const recorded = JSON.parse(readFileSync(config, 'utf8'))
    expect(recorded.python).toBe(resolved.python?.path)
    expect(recorded.libreofficeKit).toEqual({ node: process.execPath, cli: 'kept' })
  }, 60_000)

  it.runIf(requireRenderer)('loads the LibreOffice kit backend for this platform', () => {
    const kit = status?.renderers.kit
    expect(kit?.available).toBe(true)
    expect(['native', 'wasm']).toContain(kit?.backend)
    expect(status?.renderers.can_render).toEqual({ docx: true, pptx: true, xlsx: true, pdf: true })
    expect(status?.renderers.can_recalculate).toEqual({ xlsx: true })
    // The kit ships a native engine for macOS and Windows and falls back to WASM elsewhere.
    if (process.platform === 'linux')
      expect(kit?.backend).toBe('wasm')
  }, 60_000)

  it.runIf(canRecalculate)('recalculates XLSX formulas through the LibreOffice kit', () => {
    const file = create('spreadsheets', '财务 模型-计算前')
    const out = join(directory, '财务 模型-计算后.xlsx')
    const result = runCli(python!, ['xlsx', 'recalc', '--input', file, '--out', out])
    expect(result.ok).toBe(true)
    expect(result.renderer).toBe('libreoffice-kit')
    expect(['native', 'wasm']).toContain(result.engine)
    const rows = (read('xlsx', out, true) as { sheets: { rows: unknown[][] }[] }).sheets[0].rows
    expect(rows[1][3]).toBe(40)
    expect(rows[3][1]).toBe(228)
  }, 60_000)

  it.runIf(canRenderDocx)('renders DOCX pages to PNG through the LibreOffice kit', () => {
    const file = create('documents', '季度 报告-渲染')
    const out = join(directory, 'render')
    const result = runCli(python!, ['docx', 'render', '--input', file, '--out-dir', out, '--pdf'])
    expect(result.renderer).toBe('libreoffice-kit')
    expect(['native', 'wasm']).toContain(result.engine)
    expect(result.page_count).toBeGreaterThan(0)
    expect(Array.isArray(result.missing_fonts)).toBe(true)
    expect(typeof result.glyph_fallback).toBe('boolean')
    expect(result.pages.length).toBe(result.page_count)
    const pages = readdirSync(out).filter(name => name.endsWith('.png'))
    expect(pages.length).toBe(result.page_count)
    for (const page of pages)
      expect(readFileSync(join(out, page)).byteLength).toBeGreaterThan(10_000)
    const pdf = String(result.pdf)
    expect(readFileSync(pdf).subarray(0, 4).toString()).toBe('%PDF')
  }, 120_000)

  it('draws every slide in the draft preview, including pie charts', () => {
    const spec = join(directory, 'pptx-pie.json')
    const deck = join(directory, '草稿 饼图.pptx')
    writeFileSync(spec, JSON.stringify({ slides: [{ layout: 'blank', title: '分布', chart: { type: 'pie', left: 2, top: 2, width: 6, height: 4.5, categories: ['A', 'B'], series: [{ name: '占比', values: [60, 40] }] } }] }))
    expect(runCli(python!, ['pptx', 'create', '--spec', spec, '--out', deck]).ok).toBe(true)
    const result = runScript(python!, 'office-presentations', 'container_tools/render_slides.py', ['--input', deck, '--out-dir', join(directory, 'draft'), '--engine', 'draft'])
    expect(result.ok).toBe(true)
    expect(result.engine).toBe('draft-preview')
    expect(result.slides).toBe(1)
    expect(existsSync(result.files[0])).toBe(true)
  }, 60_000)

  it.runIf(canRenderDocx)('renders PPTX slides and XLSX sheets through the LibreOffice kit', () => {
    const deck = create('presentations', '渲染 检查')
    const slides = runCli(python!, ['pptx', 'render', '--input', deck, '--out-dir', join(directory, 'render-pptx')])
    expect(slides.renderer).toBe('libreoffice-kit')
    expect(slides.page_count).toBeGreaterThan(0)
    expect(slides.pages.length).toBe(slides.page_count)
    const book = create('spreadsheets', '渲染 表格')
    const sheets = runCli(python!, ['xlsx', 'render', '--input', book, '--out-dir', join(directory, 'render-xlsx')])
    expect(sheets.renderer).toBe('libreoffice-kit')
    expect(sheets.page_count).toBeGreaterThan(0)
    expect(sheets.pages.length).toBe(sheets.page_count)
    // Converted through the PDF path, so a chart anchored outside the used range
    // is rendered and verifiable instead of silently dropped.
    const pdf = String(sheets.pdf)
    expect(readFileSync(pdf).subarray(0, 4).toString()).toBe('%PDF')
    const chartText = read('pdf', pdf).pages.map((page: { text: string }) => page.text).join('\n')
    expect(chartText).toContain('收入趋势')
    for (const image of sheets.images as { page: number, sheet: string | null, range: string | null, path: string }[]) {
      expect(image.page).toBeGreaterThan(0)
      expect(image.sheet).toBeNull()
      expect(image.range).toBeNull()
      expect(readFileSync(image.path).byteLength).toBeGreaterThan(5_000)
    }
  }, 120_000)

  it.runIf(canRenderDocx)('rasterises PDF pages exported by the kit', () => {
    const file = create('documents', 'PDF 栅格-来源')
    const result = runCli(python!, ['docx', 'render', '--input', file, '--out-dir', join(directory, 'render-pdf-src'), '--pdf'])
    const pages = runCli(python!, ['pdf', 'render', '--input', String(result.pdf), '--out-dir', join(directory, 'pdf-pages')])
    expect(pages.page_count).toBeGreaterThan(0)
    expect(pages.pages.length).toBe(pages.page_count)
  }, 120_000)

  it.runIf(canRenderDocx)('reports substituted fonts instead of silently producing tofu', () => {
    const file = join(directory, '缺字体.docx')
    const script = 'import sys\nfrom docx import Document\nd = Document()\nr = d.add_paragraph("fallback 缺字体").runs[0]\nr.font.name = "NoSuchFontXYZ"\nd.save(sys.argv[1])'
    execFileSync(python!, ['-c', script, file])
    const out = join(directory, 'render-fonts')
    const result = runCli(python!, ['docx', 'render', '--input', file, '--out-dir', out])
    expect(result.missing_fonts).toContain('NoSuchFontXYZ')
    expect(result.glyph_fallback).toBe(true)
    expect((result.warnings as string[]).join('\n')).toContain('NoSuchFontXYZ')
  }, 60_000)

  it.runIf(canRenderDocx)('ignores fonts declared only by unused template styles', () => {
    // python-docx's template ships a MacroText style in Courier that no paragraph
    // uses; reporting it as missing makes every generated document look broken.
    const file = create('documents', '字体 默认')
    const result = runCli(python!, ['docx', 'render', '--input', file, '--out-dir', join(directory, 'render-default-fonts')])
    expect(result.missing_fonts).not.toContain('Courier')
    expect((result.warnings as string[]).join('\n')).not.toContain('Courier')
  }, 60_000)

  it('numbers PDF list items sequentially', () => {
    const spec = join(directory, 'pdf-numbered.json')
    const out = join(directory, '编号.pdf')
    writeFileSync(spec, JSON.stringify({ title: '编号', blocks: [{ type: 'numbered', items: ['第一项', '第二项'] }] }))
    expect(runCli(python!, ['pdf', 'create', '--spec', spec, '--out', out]).ok).toBe(true)
    const text = read('pdf', out).pages.map((page: { text: string }) => page.text).join('\n')
    expect(text).toContain('1. 第一项')
    expect(text).toContain('2. 第二项')
  })

  it('counts real headings instead of every body paragraph', () => {
    const file = create('documents', '标题 审计')
    const report = runScript(python!, 'office-documents', 'scripts/heading_audit.py', [file])
    expect(report.headings).toBe(2)
    expect(report.counts.Title).toBeUndefined()
  })

  it('adds internal navigation links without the broken default path', () => {
    const file = create('documents', '导航 基础')
    const out = join(directory, '导航 输出.docx')
    const result = runScript(python!, 'office-documents', 'scripts/internal_nav.py', [file, '--out', out])
    expect(result.ok).toBe(true)
    expect(result.entries).toBeGreaterThan(0)
    expect(read('docx', out).paragraphs.some((paragraph: { text: string }) => paragraph.text === 'Back to contents')).toBe(true)
  })

  it('rebuilds navigation idempotently instead of stacking a second block', () => {
    const file = create('documents', '导航 幂等')
    const first = join(directory, '导航 一次.docx')
    const second = join(directory, '导航 二次.docx')
    const initial = runScript(python!, 'office-documents', 'scripts/internal_nav.py', [file, '--out', first])
    const rerun = runScript(python!, 'office-documents', 'scripts/internal_nav.py', [first, '--out', second])
    expect(rerun.ok).toBe(true)
    expect(rerun.stripped).toBeGreaterThan(0)
    // The rerun must see only the real chapters, never the previous block's title.
    expect(rerun.headings.length).toBe(initial.headings.length)
    const paragraphs = read('docx', second).paragraphs.map((paragraph: { text: string }) => paragraph.text)
    expect(paragraphs.filter((text: string) => text === 'Table of Contents')).toHaveLength(1)
  })

  it('fits a table to the text width and leaves no finding', () => {
    const file = create('documents', '表格 几何')
    const out = join(directory, '表格 几何-修复.docx')
    const fixed = runScript(python!, 'office-documents', 'scripts/table_geometry.py', [file, '--fix', '--out', out])
    expect(fixed.ok).toBe(true)
    expect(fixed.findings).toEqual([])
    expect(runScript(python!, 'office-documents', 'scripts/table_geometry.py', [out]).findings).toEqual([])
  })

  it('records tracked replacements inside table cells', () => {
    const spec = join(directory, 'docx-tracked.json')
    const file = join(directory, '修订 表格.docx')
    writeFileSync(spec, JSON.stringify({ blocks: [{ type: 'table', headers: ['区域'], rows: [['EMEA'], ['APAC']] }] }))
    runCli(python!, ['docx', 'create', '--spec', spec, '--out', file])
    const tracked = join(directory, '修订 表格-跟踪.docx')
    const report = runScript(python!, 'office-documents', 'scripts/add_tracked_replacements.py', [file, '--replace', 'EMEA=Europe', '--replace', 'APAC=Asia-Pacific', '--out', tracked])
    expect(report.replacements.map((item: { paragraphs: number }) => item.paragraphs)).toEqual([1, 1])
    const accepted = join(directory, '修订 表格-接受.docx')
    expect(runScript(python!, 'office-documents', 'scripts/accept_tracked_changes.py', [tracked, '--mode', 'accept', '--out', accepted]).revisions_seen).toBe(4)
    const text = runJson(python!, ['-c', 'import json, sys; from docx import Document; d = Document(sys.argv[1]); print(json.dumps([c.text for t in d.tables for row in t.rows for c in row.cells]))', accepted])
    expect(text).toContain('Europe')
    expect(text).toContain('Asia-Pacific')
  })

  it.runIf(canRenderDocx)('renders a watermark through the header instead of writing an invisible shape', () => {
    const file = create('documents', '水印 基础')
    const marked = join(directory, '水印 标记.docx')
    const added = runScript(python!, 'office-documents', 'scripts/watermark_add.py', [file, '--text', 'CONFIDENTIAL', '--out', marked])
    expect(added.ok).toBe(true)
    expect(added.text).toBe('CONFIDENTIAL')
    expect(String(added.part)).toMatch(/header\d*\.xml$/u)
    const render = runCli(python!, ['docx', 'render', '--input', marked, '--out-dir', join(directory, 'render-watermark'), '--pdf'])
    expect(render.ok).toBe(true)
    expect(read('pdf', String(render.pdf)).pages.map((page: { text: string }) => page.text).join('').replace(/\s+/gu, '')).toContain('CONFIDENTIAL')
    const report = runScriptAudit(python!, 'office-documents', 'scripts/watermark_audit_remove.py', [marked, '--mode', 'report'])
    expect(report.status).toBe(1)
    expect(report.payload.watermarks).toHaveLength(1)
    const clean = join(directory, '水印 已移除.docx')
    expect(runScript(python!, 'office-documents', 'scripts/watermark_audit_remove.py', [marked, '--mode', 'remove', '--out', clean]).removed).toBe(1)
    expect(runScriptAudit(python!, 'office-documents', 'scripts/watermark_audit_remove.py', [clean, '--mode', 'report']).payload.watermarks).toHaveLength(0)
  }, 120_000)

  it.runIf(canRenderDocx)('merges a footnote reference together with its notes part', () => {
    const base = create('documents', '合并 基础')
    const sourceSpec = join(directory, 'docx-note.json')
    const source = join(directory, '合并 脚注源.docx')
    writeFileSync(sourceSpec, JSON.stringify({ blocks: [{ type: 'paragraph', text: '正文[[NOTE]]' }] }))
    runCli(python!, ['docx', 'create', '--spec', sourceSpec, '--out', source])
    const noted = join(directory, '合并 脚注.docx')
    expect(runScript(python!, 'office-documents', 'scripts/insert_note.py', [source, '--kind', 'footnote', '--text', '脚注内容', '--out', noted]).ok).toBe(true)
    const merged = join(directory, '合并 完成.docx')
    const report = runScript(python!, 'office-documents', 'scripts/merge_docx_append.py', [base, noted, '--out', merged])
    expect(report.notes_merged).toBe(1)
    expect(report.references_dropped).toBe(0)
    const render = runCli(python!, ['docx', 'render', '--input', merged, '--out-dir', join(directory, 'render-merge'), '--pdf'])
    expect(render.ok).toBe(true)
    const notes = runJson(python!, ['-c', 'import json, sys, zipfile; z = zipfile.ZipFile(sys.argv[1]); print(json.dumps(z.read("word/footnotes.xml").decode("utf-8")))', merged])
    expect(String(notes)).toContain('脚注内容')
  }, 120_000)
})

describe('renderer stack', () => {
  it('never shells out to PowerShell, COM, or a system office binary', () => {
    for (const relative of ['python/office_cli.py', 'src/index.ts']) {
      const source = readFileSync(join(PACKAGE_ROOT, relative), 'utf8')
      expect(source, relative).not.toMatch(/powershell|ComObject|KWps|Ket\.Application|soffice/iu)
    }
  })
})

describe.runIf(ready)('kit discovery handoff', () => {
  it('uses the path the host recorded in config.json from a tree that has no kit', () => {
    // A profile install keeps the plugin and the kit in different trees, so the
    // helper must trust the record instead of walking up.
    const kitCli = createRequire(import.meta.url).resolve('@deepseek-ai/libreoffice-kit/cli')
    const root = mkdtempSync(join(tmpdir(), 'dsh office handoff '))
    mkdirSync(join(root, 'python'), { recursive: true })
    copyFileSync(join(PACKAGE_ROOT, 'package.json'), join(root, 'package.json'))
    copyFileSync(join(PACKAGE_ROOT, 'python', 'office_cli.py'), join(root, 'python', 'office_cli.py'))
    const config = join(root, 'config.json')
    writeFileSync(config, JSON.stringify({ python: 'untouched', libreofficeKit: { node: process.execPath, cli: kitCli } }))
    const helper = join(root, 'python', 'office_cli.py')
    // Isolate the legacy config locations too: this machine may still have one, and
    // the helper reads those as fallbacks by design.
    const isolated: NodeJS.ProcessEnv = {
      ...process.env,
      DSH_OFFICE_CONFIG: config,
      APPDATA: join(root, 'appdata'),
      XDG_CONFIG_HOME: join(root, 'xdg'),
      HOME: join(root, 'home'),
    }
    delete isolated.DSH_LIBREOFFICE_KIT_CLI
    const report = JSON.parse(execFileSync(python!, [helper, 'status'], { encoding: 'utf8', env: isolated }))
    expect(report.renderers.kit.available).toBe(true)
    expect(report.renderers.kit.cli).toBe(kitCli)
    const without = JSON.parse(execFileSync(python!, [helper, 'status'], { encoding: 'utf8', env: { ...isolated, DSH_OFFICE_CONFIG: join(root, 'missing.json') } }))
    expect(without.renderers.kit.available).toBe(false)
  }, 60_000)
})
