// Contract check for the client half: drive `apply` with a recording context and
// assert the registrations the document pane needs to select this renderer.
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

const clientPath = fileURLToPath(new URL('../lib/client.js', import.meta.url))
const source = readFileSync(clientPath, 'utf8')
let record
new Function('window', source)({ __ModuleLoader__: { load: (value) => { record = value } } })
const exports = record.factory(() => ({ createElement: () => null, useState: (value) => [value, () => {}], useEffect: () => {}, useMemo: (factory) => factory() }))

let failures = 0
const check = (label, condition, detail) => {
	console.log(`${condition ? '  ok  ' : '  FAIL'} ${label}${condition || detail === undefined ? '' : ` — ${detail}`}`)
	if (!condition) failures += 1
}

const locales = []
const definitions = []
const injections = []
const registrations = []
const ctx = {
	locale: {
		bind: (namespace) => {
			locales.push(`bind:${namespace}`)
			return (key, params) => `[${namespace}.${key}${params === undefined ? '' : JSON.stringify(params)}]`
		},
		register: (namespace, dictionaries) => {
			locales.push(`register:${namespace}`)
			check(`${namespace} dictionaries have identical key sets`, JSON.stringify(Object.keys(dictionaries.zh).sort()) === JSON.stringify(Object.keys(dictionaries.en).sort()))
			return () => {}
		}
	},
	documentPreviews: {
		register: (definition) => {
			definitions.push(definition)
			return () => {}
		}
	},
	slots: {
		inject: (name, callback) => {
			injections.push(name)
			return callback()
		},
		register: (options, component) => {
			registrations.push({ options, component })
			return () => {}
		}
	},
	effect: (factory, label) => {
		const disposer = factory()
		if (typeof disposer !== 'function') throw new Error(`effect ${label} returned no disposer`)
	}
}

console.log('client contract')
exports.apply(ctx)
check('service inject list', JSON.stringify(exports.inject) === JSON.stringify(['documentPreviews', 'slots', 'locale']), JSON.stringify(exports.inject))
check('two renderer definitions', definitions.length === 2, String(definitions.length))
check('two body registrations', registrations.length === 2, String(registrations.length))
check('both bodies target the document slot', injections.every((name) => name === 'sidebar.right.tab.document'), injections.join(','))
check('locale namespace is registered and bound', locales.includes('register:officeDocumentPreview') && locales.includes('bind:officeDocumentPreview'), locales.join(','))
for (const definition of definitions) {
	check(`${definition.id} loads complete bytes`, definition.loading === 'bytes-complete')
	check(`${definition.id} outranks the builtin fallback`, definition.priority === 'external')
	check(`${definition.id} names a localized title`, typeof definition.title === 'function' && /^\[officeDocumentPreview\./u.test(definition.title()))
	check(`${definition.id} has an owning body`, registrations.some((entry) => entry.options.key === definition.id))
}
const word = definitions.find((definition) => definition.id.endsWith('/word'))
const excel = definitions.find((definition) => definition.id.endsWith('/excel'))
check('word suffixes', word.extensions.join(',') === 'docx,docm,dotx,dotm', word.extensions.join(','))
check('excel suffixes', excel.extensions.join(',') === 'xlsx,xlsm,xltx,xltm', excel.extensions.join(','))
check('bodies declare their kind', registrations.map((entry) => entry.options.inject().kind).sort().join(',') === 'excel,word')
check('bodies are components', registrations.every((entry) => typeof entry.component === 'function'))

console.log(failures === 0 ? '\nALL CHECKS PASSED' : `\n${failures} CHECK(S) FAILED`)
process.exit(failures === 0 ? 0 : 1)
