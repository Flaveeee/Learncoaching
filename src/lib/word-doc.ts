import JSZip from 'jszip'

// Erstellt und liest Word-Dateien (.docx) für den Gesprächsführer.
// Schülername, Markierungen und Notizen stehen in Inhaltssteuerelementen
// (Content Controls) mit festen Tags. Word behält diese beim Bearbeiten bei,
// so lässt sich die Datei später exakt wieder einlesen.

export type WordPhase = {
  id: string
  label: string
  title: string
  desc: string
  questions: { text: string }[]
}

export type StudentExport = {
  student: string
  duration: string
  notes: Record<string, string> // phaseId -> Notiz
  marked: Record<string, number[]> // phaseId -> Frage-Indizes
}

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml'
const TAG_PREFIX = 'lc:'

// ── Hilfsfunktionen ────────────────────────────────────────────────────

function esc(text: string) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function decodeHtml(text: string) {
  return text
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&')
}

function run(text: string, props = '') {
  const rPr = props ? `<w:rPr>${props}</w:rPr>` : ''
  return `<w:r>${rPr}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`
}

/** Wandelt Fragetext mit <strong> in Word-Runs um. */
function htmlToRuns(html: string) {
  const parts = html.split(/(<strong>[\s\S]*?<\/strong>)/g).filter(Boolean)
  return parts
    .map(part => {
      const bold = part.startsWith('<strong>')
      const text = decodeHtml(part.replace(/<[^>]+>/g, ''))
      return run(text, bold ? '<w:b/>' : '')
    })
    .join('')
}

function para(content: string, style?: string, extraPPr = '') {
  const pPr = style || extraPPr
    ? `<w:pPr>${style ? `<w:pStyle w:val="${style}"/>` : ''}${extraPPr}</w:pPr>`
    : ''
  return `<w:p>${pPr}${content}</w:p>`
}

function cell(content: string, width: number, opts: { shade?: string; vAlign?: string } = {}) {
  const shade = opts.shade ? `<w:shd w:val="clear" w:color="auto" w:fill="${opts.shade}"/>` : ''
  const vAlign = opts.vAlign ? `<w:vAlign w:val="${opts.vAlign}"/>` : ''
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${shade}${vAlign}</w:tcPr>${content}</w:tc>`
}

function table(rows: string, widths: number[]) {
  const grid = widths.map(w => `<w:gridCol w:w="${w}"/>`).join('')
  return `<w:tbl><w:tblPr><w:tblStyle w:val="LCTabelle"/><w:tblW w:w="${widths.reduce((a, b) => a + b, 0)}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`
}

function inlineSdt(tag: string, alias: string, content: string, extraPr = '') {
  return `<w:sdt><w:sdtPr><w:alias w:val="${esc(alias)}"/><w:tag w:val="${esc(TAG_PREFIX + tag)}"/>${extraPr}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`
}

function checkbox(tag: string, alias: string, checked: boolean) {
  const pr = `<w14:checkbox><w14:checked w14:val="${checked ? 1 : 0}"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>`
  const symbol = checked ? '☒' : '☐'
  const content = run(symbol, '<w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic"/><w:color w:val="2D5A3D"/><w:sz w:val="26"/>')
  return inlineSdt(tag, alias, content, pr)
}

// ── Export ─────────────────────────────────────────────────────────────

function buildDocumentXml(phases: WordPhase[], data: StudentExport) {
  const body: string[] = []
  const fullWidth = 9638

  body.push(para(run('Lerncoaching · Gesprächsprotokoll'), 'Title'))

  const infoRows = [
    ['Schüler:in', inlineSdt('student', 'Schüler:in', run(data.student, '<w:b/>'))],
    ['Datum', run(new Date().toLocaleDateString('de-DE'))],
    ['Gesprächsdauer', run(data.duration)],
  ]
    .map(([label, value]) =>
      `<w:tr>${cell(para(run(label, '<w:color w:val="5A6B5E"/>')), 2600, { shade: 'EEF3EC' })}${cell(para(value), fullWidth - 2600)}</w:tr>`,
    )
    .join('')
  body.push(table(infoRows, [2600, fullWidth - 2600]))
  body.push(para(''))

  phases.forEach((phase, pi) => {
    body.push(para(run(`Phase ${pi + 1} · ${phase.title}`), 'Heading1', pi > 0 ? '<w:pageBreakBefore/>' : ''))
    body.push(para(run(phase.desc), 'LCBeschreibung'))

    const marked = new Set(data.marked[phase.id] ?? [])
    const header = `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell(para(run('Nr.', '<w:b/><w:color w:val="FFFFFF"/>')), 700, { shade: '2D5A3D' })}${cell(para(run('Frage', '<w:b/><w:color w:val="FFFFFF"/>')), fullWidth - 1900, { shade: '2D5A3D' })}${cell(para(run('Markiert', '<w:b/><w:color w:val="FFFFFF"/>'), undefined, '<w:jc w:val="center"/>'), 1200, { shade: '2D5A3D' })}</w:tr>`
    const rows = phase.questions
      .map((q, qi) => {
        const isMarked = marked.has(qi)
        const shade = isMarked ? 'FBF1D9' : undefined
        return `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell(para(run(String(qi + 1), '<w:color w:val="5A6B5E"/>')), 700, { shade, vAlign: 'center' })}${cell(para(htmlToRuns(q.text)), fullWidth - 1900, { shade, vAlign: 'center' })}${cell(para(checkbox(`mark:${phase.id}:${qi}`, `Markierung ${phase.label} ${qi + 1}`, isMarked), undefined, '<w:jc w:val="center"/>'), 1200, { shade, vAlign: 'center' })}</w:tr>`
      })
      .join('')
    body.push(table(header + rows, [700, fullWidth - 1900, 1200]))

    body.push(para(run(`Notizen zur Phase «${phase.label}»`), 'Heading2'))
    const lines = (data.notes[phase.id] ?? '').split('\n')
    const noteParas = lines.map(line => para(line ? run(line) : '', 'LCNotiz')).join('')
    body.push(
      `<w:sdt><w:sdtPr><w:alias w:val="${esc(`Notizen ${phase.label}`)}"/><w:tag w:val="${TAG_PREFIX}note:${phase.id}"/></w:sdtPr><w:sdtContent>${noteParas}</w:sdtContent></w:sdt>`,
    )
  })

  body.push(
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`,
  )

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>${body.join('')}</w:body></w:document>`
}

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="1F2A22"/><w:lang w:val="de-DE"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:color w:val="2D5A3D"/><w:sz w:val="44"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60"/><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="2D5A3D"/></w:pBdr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:color w:val="2D5A3D"/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="320" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="2D5A3D"/><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="LCBeschreibung"><w:name w:val="LC Beschreibung"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="200"/></w:pPr><w:rPr><w:i/><w:color w:val="5A6B5E"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="LCNotiz"><w:name w:val="LC Notiz"/><w:basedOn w:val="Normal"/><w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="7A9E7E"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="F7F4EC"/><w:spacing w:after="0"/><w:ind w:left="180"/></w:pPr></w:style>
<w:style w:type="table" w:customStyle="1" w:styleId="LCTabelle"><w:name w:val="LC Tabelle"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="D5DDD3"/><w:left w:val="single" w:sz="4" w:space="0" w:color="D5DDD3"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="D5DDD3"/><w:right w:val="single" w:sz="4" w:space="0" w:color="D5DDD3"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="D5DDD3"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="D5DDD3"/></w:tblBorders><w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>`

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`

function coreXml(student: string) {
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(`Lerncoaching – ${student}`)}</dc:title><dc:creator>Lerncoaching Gesprächsführer</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`
}

export async function createStudentDocx(phases: WordPhase[], data: StudentExport): Promise<Blob> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES)
  zip.file('_rels/.rels', ROOT_RELS)
  zip.file('docProps/core.xml', coreXml(data.student))
  zip.file('word/document.xml', buildDocumentXml(phases, data))
  zip.file('word/styles.xml', STYLES_XML)
  zip.file('word/_rels/document.xml.rels', DOC_RELS)
  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
  })
}

// ── Import ─────────────────────────────────────────────────────────────

function paragraphText(p: Element) {
  let text = ''
  const walk = (node: Element) => {
    for (const child of Array.from(node.children)) {
      if (child.namespaceURI !== W_NS) continue
      if (child.localName === 't') text += child.textContent ?? ''
      else if (child.localName === 'tab') text += '\t'
      else if (child.localName === 'br' || child.localName === 'cr') text += '\n'
      else if (child.localName !== 'sdtPr' && child.localName !== 'rPr' && child.localName !== 'pPr') walk(child)
    }
  }
  walk(p)
  return text
}

function sdtTag(sdt: Element) {
  const pr = Array.from(sdt.children).find(c => c.localName === 'sdtPr')
  const tag = pr && Array.from(pr.children).find(c => c.localName === 'tag')
  return tag?.getAttributeNS(W_NS, 'val') ?? tag?.getAttribute('w:val') ?? null
}

function sdtContent(sdt: Element) {
  return Array.from(sdt.children).find(c => c.localName === 'sdtContent') ?? null
}

export async function readStudentDocx(file: Blob): Promise<StudentExport> {
  const zip = await JSZip.loadAsync(file)
  const docFile = zip.file('word/document.xml')
  if (!docFile) throw new Error('Keine Word-Datei.')
  const xml = await docFile.async('string')
  const doc = new DOMParser().parseFromString(xml, 'application/xml')

  const result: StudentExport = { student: '', duration: '', notes: {}, marked: {} }
  let found = false

  for (const sdt of Array.from(doc.getElementsByTagNameNS(W_NS, 'sdt'))) {
    const tag = sdtTag(sdt)
    if (!tag?.startsWith(TAG_PREFIX)) continue
    const key = tag.slice(TAG_PREFIX.length)
    const content = sdtContent(sdt)
    if (!content) continue

    if (key === 'student') {
      result.student = paragraphText(content).trim()
      found = true
    } else if (key.startsWith('note:')) {
      const phaseId = key.slice(5)
      const paragraphs = Array.from(content.getElementsByTagNameNS(W_NS, 'p'))
      const text = paragraphs.length > 0 ? paragraphs.map(paragraphText).join('\n') : paragraphText(content)
      result.notes[phaseId] = text.replace(/\s+$/, '')
    } else if (key.startsWith('mark:')) {
      const [, phaseId, idx] = key.split(':')
      const checkedEl = sdt.getElementsByTagNameNS(W14_NS, 'checked')[0]
      const checkedVal = checkedEl?.getAttributeNS(W14_NS, 'val') ?? checkedEl?.getAttribute('w14:val')
      const text = paragraphText(content)
      const checked = checkedVal != null ? checkedVal === '1' || checkedVal === 'true' : /[☒✓✔xX]/.test(text)
      if (checked) {
        ;(result.marked[phaseId] ??= []).push(Number(idx))
      }
    }
  }

  if (!found || !result.student) {
    throw new Error('Diese Word-Datei wurde nicht mit dem Gesprächsführer erstellt.')
  }
  return result
}
