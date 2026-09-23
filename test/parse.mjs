// Offline checks for the client bundle's parsers. Node has no `window`, no DOM
// and no React, so this harness supplies the tiny module-loader shim the bundle
// expects and then drives the exported test surface directly.
//
//   node test/parse.mjs <file.xlsx|file.docx> [...]
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

const clientPath = fileURLToPath(new URL('../lib/client.js', import.meta.url))
const source = readFileSync(clientPath, 'utf8')

let record
const window = {
	__ModuleLoader__: {
		load(value) {
			record = value
		}
	}
}
new Function('window', source)(window)
if (record === undefined) throw new Error('the bundle did not register itself')
if (record.id !== 'dsh-office-preview') throw new Error(`unexpected bundle id ${record.id}`)

const reactStub = {
	createElement: (...args) => ({ args }),
	useState: (value) => [value, () => {}],
	useEffect: () => {},
	useMemo: (factory) => factory()
}
const exports = record.factory((specifier) => {
	if (specifier === 'react') return reactStub
	throw new Error(`unexpected require(${specifier})`)
})
if (typeof exports.apply !== 'function') throw new Error('the bundle exports no apply')
if (!Array.isArray(exports.inject)) throw new Error('the bundle exports no inject array')

const t = exports.__test
let failures = 0
function check(label, condition, detail) {
	if (condition) {
		console.log(`  ok   ${label}`)
		return
	}
	failures += 1
	console.log(`  FAIL ${label}${detail === undefined ? '' : ` — ${detail}`}`)
}
function heading(text) {
	console.log(`\n${text}`)
}

//#region xml reader
heading('xml reader')
const tree = t.parseXml('<?xml version="1.0"?><a x="1 &amp; 2"><b/>text<c>t&#228;&#x4E2D;</c></a>')
const a = tree.children[0]
check('root element name', a.name === 'a')
check('attribute entity', a.attrs.x === '1 & 2', a.attrs.x)
check('self-closing sibling', a.children.filter((child) => typeof child !== 'string')[0].name === 'b')
check('text and nested text', t.textOf(a) === 'texttä中', t.textOf(a))
check('prefixed local name lookup', t.parseXml('<w:p><w:r><w:t>hi</w:t></w:r></w:p>').children[0].name === 'w:p')
//#endregion

//#region xlsx
async function checkXlsx(path) {
	heading(path)
	const bytes = new Uint8Array(readFileSync(path))
	const workbook = await t.parseXlsx(bytes)
	check('at least one sheet', workbook.sheets.length > 0)
	console.log(`  sheets: ${workbook.sheets.map((sheet) => sheet.name).join(' | ')}`)
	for (const sheet of workbook.sheets) {
		const rows = sheet.rows.length
		const columns = Math.max(...sheet.rows.map((row) => row.filter(Boolean).length), 0)
		console.log(`  ${sheet.name}: ${rows} rows × ${columns} used columns, ${sheet.merges.length} merges`)
	}
	const first = workbook.sheets[0]
	const head = first.rows.slice(0, 8).map((row) => (row ?? []).map((cell) => (cell === undefined ? '' : cell.text)).join(' | '))
	console.log(head.map((line) => `    ${line}`).join('\n'))
	const flat = first.rows.flat().filter(Boolean).map((cell) => cell.text)
	check('non-empty grid', flat.length > 0)
	const dates = flat.filter((text) => /^\d{4}-\d{2}-\d{2}/u.test(text))
	check('date cells render as dates', dates.length > 0, `found ${dates.length}: ${dates.slice(0, 3).join(', ')}`)
}
//#endregion

//#region docx
async function checkDocx(path) {
	heading(path)
	const parsed = await t.parseDocx(new Uint8Array(readFileSync(path)))
	check('document has blocks', parsed.blocks.length > 0)
	check('no empty runs on headings', parsed.blocks.every((block) => block.kind !== 'p' || Array.isArray(block.runs)))
	const tables = parsed.blocks.filter((block) => block.kind === 'table')
	console.log(`  ${parsed.blocks.length} top-level blocks, ${tables.length} tables`)
	for (const block of parsed.blocks.slice(0, 20)) {
		if (block.kind === 'table') {
			console.log(`    [table ${block.rows.length}×${block.rows[0]?.length ?? 0}]`)
			for (const row of block.rows.slice(0, 3)) {
				console.log(`      ${row.map((cell) => cell.blocks.map((nested) => (nested.kind === 'table' ? '[table]' : nested.runs.map((run) => run.text).join(''))).join('')).join(' | ')}`)
			}
			continue
		}
		const text = block.runs.map((run) => run.text).join('')
		console.log(`    ${block.level > 0 ? `h${block.level} ` : ''}${block.bullet}${text}`)
	}
}
//#endregion

const samples = process.argv.slice(2)
if (samples.length === 0) throw new Error('pass at least one .xlsx or .docx path')
for (const path of samples) {
	if (path.toLowerCase().endsWith('.xlsx')) await checkXlsx(path)
	else if (path.toLowerCase().endsWith('.docx')) await checkDocx(path)
	else throw new Error(`unsupported sample ${path}`)
}

console.log(failures === 0 ? '\nALL CHECKS PASSED' : `\n${failures} CHECK(S) FAILED`)
process.exit(failures === 0 ? 0 : 1)
