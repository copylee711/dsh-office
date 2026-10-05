/**
 * LaTeX formulas to OMML, the equation format of .docx files.
 *
 * Microsoft Word builds its equations itself, from the linear text the helper writes; WPS Office does not. For WPS
 * the formulas are converted here — temml reads the LaTeX into a MathML tree, and that tree is written out as OMML —
 * and the helper inserts the result as it is.
 */
import temml from 'temml'

interface MathNode {
  type?: string
  text?: string
  attributes?: Record<string, string>
  classes?: string[]
  children?: MathNode[]
}

const render = (temml as unknown as { __renderToMathMLTree(tex: string, options: Record<string, unknown>): MathNode }).__renderToMathMLTree

const NARY = new Set(['∑', '∏', '∐', '∫', '∬', '∭', '∮', '∯', '∰', '⋃', '⋂', '⋁', '⋀', '⨁', '⨂', '⨀'])
/** What MathML writes over a letter, and the combining mark OMML wants for it. */
const ACCENTS: Record<string, string> = {
  'ˆ': '̂', '^': '̂', '˜': '̃', '~': '̃', '→': '⃗', '˙': '̇', '¨': '̈',
  'ˇ': '̌', '˘': '̆', '´': '́', '`': '̀', '⃗': '⃗', '˚': '̊', '←': '⃖', '↔': '⃡',
}
const BARS = new Set(['‾', '¯', '―', '_', '̲'])

const esc = (text: string): string => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

const FONT = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>'

function run(text: string, style: '' | 'plain' | 'normal' = ''): string {
  if (text.length === 0) return ''
  const props = style === 'normal' ? '<m:rPr><m:nor/></m:rPr>' : style === 'plain' ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : ''
  return `<m:r>${props}${FONT}<m:t xml:space="preserve">${esc(text)}</m:t></m:r>`
}

const textOf = (node: MathNode | undefined): string => node === undefined ? '' : node.text !== undefined && node.type === undefined ? node.text : (node.children ?? []).map(textOf).join('')

const isToken = (node: MathNode | undefined, type: string): boolean => node !== undefined && node.type === type

/** The children of a node as one sequence: rows and the like that only group are opened up. */
function flat(nodes: MathNode[]): MathNode[] {
  const out: MathNode[] = []
  for (const node of nodes) {
    if (node.type === undefined && node.children !== undefined) out.push(...flat(node.children))
    else if ((node.type === 'mrow' || node.type === 'mstyle' || node.type === 'mpadded' || node.type === 'semantics') && !isFenced(node)) out.push(...flat(node.children ?? []))
    else out.push(node)
  }
  return out
}

function isFenced(node: MathNode): boolean {
  const children = node.children ?? []
  if (node.type !== 'mrow' || children.length < 2) return false
  const first = children[0]!, last = children[children.length - 1]!
  return isToken(first, 'mo') && first.attributes?.fence === 'true' && isToken(last, 'mo') && last.attributes?.fence === 'true'
}

const naryOf = (node: MathNode): string | undefined => {
  if (!['msubsup', 'munderover', 'msub', 'msup', 'munder', 'mover'].includes(node.type ?? '')) return undefined
  const base = node.children?.[0]
  const sign = isToken(base, 'mo') ? textOf(base) : ''
  return NARY.has(sign) ? sign : undefined
}

function space(node: MathNode): string {
  const width = /^(-?[\d.]+)em$/.exec(node.attributes?.width ?? '')
  const em = width ? Number(width[1]) : 0
  if (em >= 1.5) return run('  ')
  if (em >= 0.8) return run(' ')
  if (em >= 0.2) return run(' ')
  if (em >= 0.1) return run(' ')
  return ''
}

function table(node: MathNode): string {
  const rows = (node.children ?? []).filter(row => row.type === 'mtr' || row.type === 'mlabeledtr')
  const columns = Math.max(1, ...rows.map(row => (row.children ?? []).length))
  const side: string[] = []
  for (let c = 0; c < columns; c++) {
    const cell = rows.map(row => row.children?.[c]).find(Boolean)
    side.push(cell?.classes?.includes('tml-left') ? 'left' : cell?.classes?.includes('tml-right') ? 'right' : 'center')
  }
  const mcs = side.map(jc => `<m:mc><m:mcPr><m:count m:val="1"/><m:mcJc m:val="${jc}"/></m:mcPr></m:mc>`).join('')
  const body = rows.map(row => {
    const cells = (row.children ?? []).map(cell => `<m:e>${seq(cell.children ?? [])}</m:e>`)
    while (cells.length < columns) cells.push('<m:e/>')
    return `<m:mr>${cells.join('')}</m:mr>`
  }).join('')
  return `<m:m><m:mPr><m:mcs>${mcs}</m:mcs></m:mPr>${body}</m:m>`
}

const arg = (node: MathNode | undefined): string => node === undefined ? '' : seq([node])

function one(node: MathNode): string {
  const kids = node.children ?? []
  switch (node.type) {
    case 'mi': {
      const text = textOf(node)
      return run(text, node.attributes?.mathvariant === 'normal' || [...text].length > 1 ? 'plain' : '')
    }
    case 'mn': return run(textOf(node))
    case 'mo': {
      const text = textOf(node)
      // The invisible "function application" and "times" marks have no place in OMML.
      return text === '⁡' || text === '⁢' || text === '⁣' ? '' : run(text)
    }
    case 'mtext': return run(textOf(node), 'normal')
    case 'mspace': return space(node)
    case 'mfrac': {
      const bar = node.attributes?.linethickness === '0px' || node.attributes?.linethickness === '0' ? '<m:fPr><m:type m:val="noBar"/></m:fPr>' : ''
      return `<m:f>${bar}<m:num>${arg(kids[0])}</m:num><m:den>${arg(kids[1])}</m:den></m:f>`
    }
    case 'msqrt': return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${seq(kids)}</m:e></m:rad>`
    case 'mroot': return `<m:rad><m:deg>${arg(kids[1])}</m:deg><m:e>${arg(kids[0])}</m:e></m:rad>`
    case 'msup': return `<m:sSup><m:e>${arg(kids[0])}</m:e><m:sup>${arg(kids[1])}</m:sup></m:sSup>`
    case 'msub': return `<m:sSub><m:e>${arg(kids[0])}</m:e><m:sub>${arg(kids[1])}</m:sub></m:sSub>`
    case 'msubsup': return `<m:sSubSup><m:e>${arg(kids[0])}</m:e><m:sub>${arg(kids[1])}</m:sub><m:sup>${arg(kids[2])}</m:sup></m:sSubSup>`
    case 'mmultiscripts': return `<m:sPre><m:sub>${arg(kids[3])}</m:sub><m:sup>${arg(kids[4])}</m:sup><m:e>${arg(kids[0])}</m:e></m:sPre>`
    case 'mover': {
      const mark = textOf(kids[1])
      if (isToken(kids[1], 'mo') && BARS.has(mark)) return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>${arg(kids[0])}</m:e></m:bar>`
      if (isToken(kids[1], 'mo') && ACCENTS[mark] !== undefined) return `<m:acc><m:accPr><m:chr m:val="${ACCENTS[mark]}"/></m:accPr><m:e>${arg(kids[0])}</m:e></m:acc>`
      if (isToken(kids[1], 'mo') && mark === '⏞') return `<m:groupChr><m:groupChrPr><m:chr m:val="⏞"/><m:pos m:val="top"/><m:vertJc m:val="bot"/></m:groupChrPr><m:e>${arg(kids[0])}</m:e></m:groupChr>`
      return `<m:limUpp><m:e>${arg(kids[0])}</m:e><m:lim>${arg(kids[1])}</m:lim></m:limUpp>`
    }
    case 'munder': {
      const mark = textOf(kids[1])
      if (isToken(kids[1], 'mo') && BARS.has(mark)) return `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr><m:e>${arg(kids[0])}</m:e></m:bar>`
      if (isToken(kids[1], 'mo') && mark === '⏟') return `<m:groupChr><m:groupChrPr><m:chr m:val="⏟"/><m:pos m:val="bot"/><m:vertJc m:val="top"/></m:groupChrPr><m:e>${arg(kids[0])}</m:e></m:groupChr>`
      return `<m:limLow><m:e>${arg(kids[0])}</m:e><m:lim>${arg(kids[1])}</m:lim></m:limLow>`
    }
    case 'munderover': return `<m:limUpp><m:e><m:limLow><m:e>${arg(kids[0])}</m:e><m:lim>${arg(kids[1])}</m:lim></m:limLow></m:e><m:lim>${arg(kids[2])}</m:lim></m:limUpp>`
    case 'menclose': {
      const notation = node.attributes?.notation ?? ''
      if (notation.includes('top')) return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>${seq(kids)}</m:e></m:bar>`
      if (notation.includes('bottom')) return `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr><m:e>${seq(kids)}</m:e></m:bar>`
      if (notation.includes('box')) return `<m:borderBox><m:e>${seq(kids)}</m:e></m:borderBox>`
      return seq(kids)
    }
    case 'mtable': return table(node)
    case 'mphantom': return ''
    case 'mrow': {
      if (isFenced(node)) {
        const open = textOf(kids[0]), close = textOf(kids[kids.length - 1])
        return `<m:d><m:dPr><m:begChr m:val="${esc(open)}"/><m:endChr m:val="${esc(close)}"/></m:dPr><m:e>${seq(kids.slice(1, -1))}</m:e></m:d>`
      }
      return seq(kids)
    }
    default: return seq(kids)
  }
}

/** A run of sibling nodes. A sum, product or integral takes the node after it as what it applies to. */
function seq(nodes: MathNode[]): string {
  const list = flat(nodes)
  let out = ''
  for (let i = 0; i < list.length; i++) {
    const node = list[i]!
    const sign = naryOf(node)
    if (sign === undefined) { out += one(node); continue }
    const kids = node.children ?? []
    const under = node.type === 'munderover' || node.type === 'munder' || node.type === 'mover'
    const sub = node.type === 'msup' || node.type === 'mover' ? undefined : kids[1]
    const sup = node.type === 'msup' || node.type === 'mover' ? kids[1] : node.type === 'msubsup' || node.type === 'munderover' ? kids[2] : undefined
    // What follows, up to the next sign that sets it apart (=, +, a comma ..), is the operand; a space alone is not.
    let operand = ''
    while (i + 1 < list.length && list[i + 1]!.type === 'mspace') i++
    if (i + 1 < list.length && !(isToken(list[i + 1], 'mo') && !isFenced(list[i + 1]!))) operand = one(list[++i]!)
    out += `<m:nary><m:naryPr><m:chr m:val="${sign}"/><m:limLoc m:val="${under ? 'undOvr' : 'subSup'}"/>${sub === undefined ? '<m:subHide m:val="1"/>' : ''}${sup === undefined ? '<m:supHide m:val="1"/>' : ''}</m:naryPr><m:sub>${arg(sub)}</m:sub><m:sup>${arg(sup)}</m:sup><m:e>${operand}</m:e></m:nary>`
  }
  return out
}

/** One formula as an `<m:oMath>` element, or undefined when the LaTeX cannot be read. */
export function latexToOmml(tex: string, display: boolean): string | undefined {
  try {
    const tree = render(tex, { displayMode: display, annotate: false, throwOnError: true, strict: false })
    const body = seq(tree.children ?? [])
    return body.length === 0 ? undefined : `<m:oMath>${body}</m:oMath>`
  } catch {
    return undefined
  }
}

const DOLLARS = /\$\$([^$\r\n]+?)\$\$|\$([^$\r\n]+?)\$(?!\$)/g
const TAG = /\\tag\s*\{([^}]*)\}/

/**
 * Every formula written between dollar signs anywhere in a batch of operations, converted: keyed "d:" + source for
 * one on a line of its own and "i:" + source for one inside a line, the source as the helper will read it.
 */
export function formulasOf(value: unknown, found: Record<string, string> = {}): Record<string, string> {
  if (typeof value === 'string') {
    if (!value.includes('$')) return found
    for (const match of value.matchAll(DOLLARS)) {
      const display = match[1] !== undefined
      const source = (match[1] ?? match[2] ?? '').trim().replace(TAG, '').trim()
      const key = (display ? 'd:' : 'i:') + source
      if (source.length === 0 || found[key] !== undefined) continue
      const omml = latexToOmml(source, display)
      if (omml !== undefined) found[key] = omml
    }
  } else if (Array.isArray(value)) {
    for (const item of value) formulasOf(item, found)
  } else if (value !== null && typeof value === 'object') {
    for (const [name, item] of Object.entries(value)) {
      // A formula in a field of its own (the cards of a formula slide) is written without dollar signs.
      if (name === 'latex' && typeof item === 'string') {
        const source = item.trim().replace(/^\$+|\$+$/g, '').trim()
        const omml = source.length === 0 || found[`d:${source}`] !== undefined ? undefined : latexToOmml(source, true)
        if (omml !== undefined) found[`d:${source}`] = omml
      } else formulasOf(item, found)
    }
  }
  return found
}
