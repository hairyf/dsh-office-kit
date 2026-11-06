import { execFileSync } from 'node:child_process'
import { existsSync, readFileSync } from 'node:fs'
import process from 'node:process'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { officePaths, venvPython } from '../src/python'

function configuredPython(): string | undefined {
  // An explicitly supplied runtime must fail visibly if it cannot run the test.
  if (process.env.DSH_OFFICE_PYTHON)
    return process.env.DSH_OFFICE_PYTHON
  const paths = officePaths()
  for (const config of [paths.configFile, ...paths.legacyConfigFiles]) {
    if (!existsSync(config))
      continue
    const parsed = JSON.parse(readFileSync(config, 'utf8')) as { python?: unknown }
    if (typeof parsed.python === 'string' && existsSync(parsed.python))
      return parsed.python
  }
  const python = venvPython(paths.venv)
  return existsSync(python) ? python : undefined
}

const python = configuredPython()
const script = fileURLToPath(new URL('../assets/skills/office-documents/scripts/replace_text.py', import.meta.url))
const setup = String.raw`
import importlib.util
import json
import sys
from io import BytesIO
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

spec = importlib.util.spec_from_file_location("replace_text", sys.argv[1])
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)

def styled_run(paragraph, text, font, east_asia=None):
    run = paragraph.add_run(text)
    run.font.name = font
    if east_asia:
        run._r.rPr.rFonts.set(qn("w:eastAsia"), east_asia)
    return run

def assert_preserved(paragraph, nodes, properties):
    assert module.paragraph_text_nodes(paragraph) == nodes, "original text nodes changed"
    assert [etree.tostring(rpr) for rpr in paragraph._p.xpath(".//w:rPr")] == properties, "run properties changed"
`

describe.skipIf(python == null)('dOCX formatting-preserving replacement', () => {
  it('keeps the numeric run font when replacing an entire mixed-font paragraph', () => {
    const regression = String.raw`
paragraph = Document().add_paragraph()
styled_run(paragraph, "预算 ", "Calibri", "宋体")
styled_run(paragraph, "18%", "Times New Roman")
styled_run(paragraph, " 增长", "Calibri", "宋体")
nodes = module.paragraph_text_nodes(paragraph)
properties = [etree.tostring(rpr) for rpr in paragraph._p.xpath(".//w:rPr")]
assert module.replace_in_paragraph(paragraph, "预算 18% 增长", "预算 19% 增长") == 1
assert [node.text for node in nodes] == ["预算 ", "19%", " 增长"], [node.text for node in nodes]
assert_preserved(paragraph, nodes, properties)
print("ok")
`
    expect(runRegression(regression)).toBe('ok')
  })

  it.each([
    ['longer', '12AXB34', ['prefix 12', 'AXB', '34 suffix']],
    ['shorter', '12A34', ['prefix 12', 'A', '34 suffix']],
    ['deleted', '', ['prefix ', '', ' suffix']],
  ])('preserves surrounding formatting when the replacement is %s', (_, replacement, expected) => {
    const regression = String.raw`
paragraph = Document().add_paragraph()
styled_run(paragraph, "prefix 12", "Arial").bold = True
styled_run(paragraph, "AB", "Times New Roman").italic = True
styled_run(paragraph, "34 suffix", "Courier New").underline = True
nodes = module.paragraph_text_nodes(paragraph)
properties = [etree.tostring(rpr) for rpr in paragraph._p.xpath(".//w:rPr")]
replacement, expected = json.loads(sys.argv[2])
assert module.replace_in_paragraph(paragraph, "12AB34", replacement) == 1
assert [node.text or "" for node in nodes] == expected, [node.text for node in nodes]
assert_preserved(paragraph, nodes, properties)
print("ok")
`
    expect(runRegression(regression, [replacement, expected])).toBe('ok')
  })

  it('counts contained and cross-run matches once without processing inserted text', () => {
    const regression = String.raw`
paragraph = Document().add_paragraph()
styled_run(paragraph, "prefix ", "Arial").bold = True
styled_run(paragraph, "ab gap a", "Times New Roman").italic = True
styled_run(paragraph, "b suffix", "Courier New").underline = True
nodes = module.paragraph_text_nodes(paragraph)
properties = [etree.tostring(rpr) for rpr in paragraph._p.xpath(".//w:rPr")]
assert module.replace_in_paragraph(paragraph, "ab", "ab ab") == 2
assert "".join(node.text or "" for node in nodes) == "prefix ab ab gap ab ab suffix"
assert nodes[0].text == "prefix "
assert nodes[-1].text.endswith(" suffix")
assert_preserved(paragraph, nodes, properties)
assert module.replace_in_paragraph(paragraph, "missing", "unused") == 0
assert module.replace_in_paragraph(paragraph, "", "unused") == 0
print("ok")
`
    expect(runRegression(regression)).toBe('ok')
  })

  it('keeps hyperlinks and marks newly exposed whitespace for preservation', () => {
    const regression = String.raw`
document = Document()
paragraph = document.add_paragraph()
styled_run(paragraph, "head A", "Arial").bold = True
linked = styled_run(paragraph, "B", "Times New Roman")
linked.italic = True
hyperlink = OxmlElement("w:hyperlink")
hyperlink.set(qn("w:anchor"), "bookmark")
paragraph._p.append(hyperlink)
hyperlink.append(linked._r)
styled_run(paragraph, "X tail", "Courier New").underline = True
nodes = module.paragraph_text_nodes(paragraph)
properties = [etree.tostring(rpr) for rpr in paragraph._p.xpath(".//w:rPr")]
assert module.replace_in_paragraph(paragraph, "ABX", " C ") == 1
assert [node.text or "" for node in nodes] == ["head  C ", "", " tail"]
assert paragraph._p.xpath(".//w:hyperlink") == [hyperlink]
assert hyperlink.get(qn("w:anchor")) == "bookmark"
assert_preserved(paragraph, nodes, properties)
assert nodes[0].get(qn("xml:space")) == "preserve"
assert nodes[2].get(qn("xml:space")) == "preserve"
output = BytesIO()
document.save(output)
output.seek(0)
reloaded = Document(output).paragraphs[0]
assert [node.text or "" for node in module.paragraph_text_nodes(reloaded)] == ["head  C ", "", " tail"]
assert reloaded._p.xpath(".//w:hyperlink")[0].get(qn("w:anchor")) == "bookmark"
print("ok")
`
    expect(runRegression(regression)).toBe('ok')
  })
})

function runRegression(regression: string, values?: unknown): string {
  return execFileSync(python!, ['-c', setup + regression, script, JSON.stringify(values ?? null)], {
    encoding: 'utf8',
    env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONDONTWRITEBYTECODE: '1' },
  }).trim()
}
