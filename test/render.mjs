// Render-path check: run the component tree the pane would mount, with a tiny
// element shim standing in for React, so a render-time crash (an empty table, a
// merge past the last column, a title row without cells) fails here rather than
// in the user's sidebar.
//
//   node test/render.mjs <file.xlsx|file.docx> [...]
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

const clientPath = fileURLToPath(new URL('../lib/client.js', import.meta.url))
let record
new Function('window', readFileSync(clientPath, 'utf8'))({ __ModuleLoader__: { load: (value) => { record = value } } })

const element = (type, props, ...children) => ({ type, props: props ?? {}, children: children.flat() })
const reactStub = {
	createElement: element,
	useState: (value) => [value, () => {}],
	useEffect: () => {},
	useMemo: (factory) => factory()
}
const exports = record.factory(() => reactStub)
const t = exports.__test

let failures = 0
const check = (label, condition, detail) => {
	console.log(`${condition ? '  ok  ' : '  FAIL'} ${label}${condition || detail === undefined ? '' : ` — ${detail}`}`)
	if (!condition) failures += 1
}

/** Expand a produced element tree into plain text lines, invoking components. */
function render(node) {
	if (node === null || node === undefined || node === false || node === true) return ''
	if (typeof node === 'string' || typeof node === 'number') return String(node)
	if (Array.isArray(node)) return node.map(render).join('')
	if (typeof node.type === 'function') return render(node.type({ ...node.props, children: node.children }))
	return node.children.map(render).join('')
}

const messages = { loading: 'loading', empty: 'empty', failed: 'failed', hint: 'hint' }
const locale = (key) => messages[key] ?? key

for (const path of process.argv.slice(2)) {
	console.log(`\n${path}`)
	const bytes = new Uint8Array(readFileSync(path))
	const lower = path.toLowerCase()
	if (lower.endsWith('.xlsx')) {
		const workbook = await t.parseXlsx(bytes)
		const body = t.WorkbookBody({ sheets: workbook.sheets, t: locale })
		const text = render(body)
		check('workbook renders text', text.trim().length > 0)
		check('sheet names appear', text.includes(workbook.sheets[0].name), workbook.sheets[0].name)
		check('cells appear', text.includes('1') && text.length > 40)
		console.log(`    ${text.replace(/\s+/gu, ' ').slice(0, 160)}`)
		// A single-cell sheet, an empty sheet and a merge past the last column are
		// the shapes that crash a naive grid.
		const edge = {
			name: 'edge',
			rows: [[{ text: 'only' }], [], undefined],
			merges: [{ row: 0, column: 3, rowSpan: 2, colSpan: 4 }],
			widths: []
		}
		check('edge-case grid renders', render(t.SheetGrid({ sheet: { ...edge, rows: edge.rows.filter(Boolean) }, t: locale })).includes('only'))
		check('empty grid renders', typeof render(t.SheetGrid({ sheet: { name: 'empty', rows: [], merges: [], widths: [] }, t: locale })) === 'string')
	} else if (lower.endsWith('.docx')) {
		const parsed = await t.parseDocx(bytes)
		const body = t.WordBody({ blocks: parsed.blocks })
		const text = render(body)
		check('document renders text', text.trim().length > 0)
		console.log(`    ${text.replace(/\s+/gu, ' ').slice(0, 160)}`)
		check('empty document renders', t.WordBody({ blocks: [] }) !== undefined)
		check('empty table renders', t.WordBody({ blocks: [{ kind: 'table', rows: [] }] }) !== undefined)
	}
}

console.log(failures === 0 ? '\nALL CHECKS PASSED' : `\n${failures} CHECK(S) FAILED`)
process.exit(failures === 0 ? 0 : 1)
