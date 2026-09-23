window.__ModuleLoader__.load({
	id: "dsh-office-preview",
	factory: (require) => {
		var module = { exports: {} };
		var exports = module.exports;
		Object.defineProperty(exports, Symbol.toStringTag, { value: "Module" });
		const react = require("react");
		//#region src/client/style.css
		/**
		 * The pane's own styles. Every renderer here owns its layout, so the rules
		 * reset the document pane's monospace/`white-space: pre` body into a normal
		 * flowing document and follow the Harness alias tokens so light and dark
		 * themes need no second definition.
		 */
		const css = [
			".dso-root{box-sizing:border-box;min-width:100%;min-height:100%;margin:0;padding:10px 12px 28px;white-space:normal;color:var(--dsw-alias-label-primary);font-family:var(--dsw-font-family,system-ui,sans-serif);font-size:13px;line-height:1.65}",
			".dso-status{box-sizing:border-box;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;min-height:100%;padding:24px 16px;color:var(--dsw-alias-label-secondary);font-size:13px;text-align:center;white-space:normal}",
			".dso-status-name{color:var(--dsw-alias-label-primary)}",
			".dso-note{color:var(--dsw-alias-label-tertiary);font-size:12px}",
			".dso-body{max-width:900px;margin:0 auto}",
			".dso-p{margin:0 0 .55em;white-space:pre-wrap;word-break:break-word}",
			".dso-h{font-weight:600;line-height:1.35;margin:.9em 0 .45em}",
			".dso-list{margin:0 0 .5em;white-space:pre-wrap}",
			".dso-table{border-collapse:collapse;margin:.35em 0 1em;font-size:12.5px}",
			".dso-table td,.dso-table th{border:.5px solid var(--dsw-alias-border-l3);padding:4px 8px;vertical-align:top;white-space:pre-wrap;word-break:break-word}",
			".dso-sheetbar{position:sticky;top:0;z-index:2;display:flex;flex-wrap:wrap;gap:4px;margin:0 0 8px;padding:6px 0;background:var(--dsw-alias-bg-layer-1,transparent)}",
			".dso-sheet{appearance:none;border:.5px solid var(--dsw-alias-border-l3);border-radius:6px;background:var(--dsw-alias-bg-layer-2,transparent);color:var(--dsw-alias-label-secondary);cursor:pointer;font:inherit;font-size:12px;padding:3px 10px}",
			".dso-sheet[data-active=true]{background:var(--dsw-alias-interactive-bg-hover);color:var(--dsw-alias-label-primary);border-color:var(--dsw-alias-border-l2)}",
			".dso-sheetwrap{overflow:auto;max-width:100%}",
			".dso-grid{border-collapse:separate;border-spacing:0;font-variant-numeric:tabular-nums;font-size:12.5px}",
			".dso-grid td,.dso-grid th{border-right:.5px solid var(--dsw-alias-border-l3);border-bottom:.5px solid var(--dsw-alias-border-l3);padding:3px 7px;min-width:52px;max-width:420px;white-space:pre-wrap;word-break:break-word;vertical-align:top}",
			".dso-grid tr:first-child td{border-top:.5px solid var(--dsw-alias-border-l3)}",
			".dso-grid td:first-child{border-left:.5px solid var(--dsw-alias-border-l3)}",
			".dso-grid thead th{position:sticky;top:0;z-index:1;background:var(--dsw-alias-bg-layer-2,#1c1c1c);color:var(--dsw-alias-label-tertiary);font-weight:400;text-align:center;font-size:11px}",
			".dso-rowhead{position:sticky;left:0;z-index:1;background:var(--dsw-alias-bg-layer-2,#1c1c1c);color:var(--dsw-alias-label-tertiary);font-weight:400;text-align:right;font-size:11px;min-width:34px}",
			".dso-corner{z-index:3;left:0}",
			".dso-b{font-weight:600}",
			".dso-c{text-align:center}",
			".dso-r{text-align:right}",
			".dso-rowwrap{background:var(--dsw-alias-bg-layer-2,rgba(127,127,127,.08))}"
		].join("");
		const cssTagId = "dsh-office-preview/OfficePreview.css";
		if (typeof document !== "undefined" && document.querySelector("style[data-plugin-css=" + JSON.stringify(cssTagId) + "]") === null) {
			const tag = document.createElement("style");
			tag.dataset.plugin = "dsh-office-preview";
			tag.dataset.pluginCss = cssTagId;
			tag.textContent = css;
			document.head.appendChild(tag);
		}
		//#endregion
		//#region src/client/xml.js
		/**
		 * A purpose-built XML reader for Office Open XML parts.
		 *
		 * OOXML is a small, regular subset: no DOCTYPE, no entity definitions
		 * beyond the five predefined ones, no mixed-content elements that matter
		 * here. A DOM would therefore cost more than it returns — and the browser's
		 * `DOMParser` is unavailable in the Node test harness — so this reader
		 * builds the same shallow tree with explicit `w:`/`r:` prefixed names kept
		 * intact, and matching ignores the prefix.
		 */
		/** Decode the predefined entities and numeric character references. */
		function decodeEntities(text) {
			if (text.indexOf("&") === -1) return text;
			return text.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z][a-zA-Z0-9]*);/gu, (match, body) => {
				if (body.charAt(0) === "#") {
					const code = body.charAt(1) === "x" || body.charAt(1) === "X" ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
					if (!Number.isFinite(code)) return match;
					try {
						return String.fromCodePoint(code);
					} catch {
						return match;
					}
				}
				switch (body) {
					case "amp": return "&";
					case "lt": return "<";
					case "gt": return ">";
					case "quot": return "\"";
					case "apos": return "'";
					default: return match;
				}
			});
		}
		/** Index of the `>` that closes a tag, ignoring quoted attribute values. */
		function tagEnd(source, start) {
			let quote = "";
			for (let index = start + 1; index < source.length; index += 1) {
				const character = source.charAt(index);
				if (quote !== "") {
					if (character === quote) quote = "";
					continue;
				}
				if (character === "\"" || character === "'") {
					quote = character;
					continue;
				}
				if (character === ">") return index;
			}
			return -1;
		}
		/** Read `name="value"` pairs that start at `from` into `into`. */
		function readAttributes(source, from, into) {
			let index = from;
			while (index < source.length) {
				while (index < source.length && /\s/u.test(source.charAt(index))) index += 1;
				if (index >= source.length) return;
				let start = index;
				while (index < source.length && !/[\s=]/u.test(source.charAt(index))) index += 1;
				const name = source.slice(start, index);
				while (index < source.length && /\s/u.test(source.charAt(index))) index += 1;
				if (source.charAt(index) !== "=") {
					if (name !== "") into[name] = "";
					continue;
				}
				index += 1;
				while (index < source.length && /\s/u.test(source.charAt(index))) index += 1;
				const quote = source.charAt(index);
				if (quote === "\"" || quote === "'") {
					const end = source.indexOf(quote, index + 1);
					const value = source.slice(index + 1, end === -1 ? source.length : end);
					if (name !== "") into[name] = decodeEntities(value);
					index = end === -1 ? source.length : end + 1;
				} else {
					start = index;
					while (index < source.length && !/\s/u.test(source.charAt(index))) index += 1;
					if (name !== "") into[name] = decodeEntities(source.slice(start, index));
				}
			}
		}
		/**
		 * Parse one XML document into `{ name, attrs, children }` nodes.
		 * @param source - XML text.
		 * @returns the `#document` root whose children carry the tree.
		 */
		function parseXml(source) {
			const root = { name: "#document", attrs: Object.create(null), children: [] };
			const stack = [root];
			const top = () => stack[stack.length - 1];
			const appendText = (text) => {
				if (text !== "") top().children.push(decodeEntities(text));
			};
			let index = 0;
			const length = source.length;
			while (index < length) {
				const open = source.indexOf("<", index);
				if (open === -1) {
					appendText(source.slice(index));
					break;
				}
				if (open > index) appendText(source.slice(index, open));
				if (source.startsWith("<!--", open)) {
					const end = source.indexOf("-->", open + 4);
					index = end === -1 ? length : end + 3;
					continue;
				}
				if (source.startsWith("<![CDATA[", open)) {
					const end = source.indexOf("]]>", open + 9);
					top().children.push(source.slice(open + 9, end === -1 ? length : end));
					index = end === -1 ? length : end + 3;
					continue;
				}
				if (source.startsWith("<?", open)) {
					const end = source.indexOf("?>", open + 2);
					index = end === -1 ? length : end + 2;
					continue;
				}
				if (source.startsWith("<!", open)) {
					const end = source.indexOf(">", open + 2);
					index = end === -1 ? length : end + 1;
					continue;
				}
				const close = tagEnd(source, open);
				if (close === -1) {
					appendText(source.slice(open));
					break;
				}
				const raw = source.slice(open + 1, close);
				index = close + 1;
				if (raw.charAt(0) === "/") {
					const name = raw.slice(1).trim();
					for (let depth = stack.length - 1; depth > 0; depth -= 1) {
						if (stack[depth].name === name) {
							stack.length = depth;
							break;
						}
					}
					continue;
				}
				const selfClosing = raw.endsWith("/");
				const body = selfClosing ? raw.slice(0, -1) : raw;
				const space = body.search(/\s/u);
				const name = space === -1 ? body : body.slice(0, space);
				const node = { name, attrs: Object.create(null), children: [] };
				readAttributes(body, space === -1 ? body.length : space, node.attrs);
				if (name !== "") top().children.push(node);
				if (!selfClosing) stack.push(node);
			}
			return root;
		}
		/** The element name without its namespace prefix. */
		function localOf(name) {
			const colon = name.lastIndexOf(":");
			return colon === -1 ? name : name.slice(colon + 1);
		}
		/** Element children of a node, in document order. */
		function elementChildren(node) {
			return node.children.filter((child) => typeof child !== "string");
		}
		/** Direct children whose local name matches. */
		function childrenNamed(node, name) {
			return elementChildren(node).filter((child) => localOf(child.name) === name);
		}
		/** The first direct child whose local name matches. */
		function childNamed(node, name) {
			return elementChildren(node).find((child) => localOf(child.name) === name);
		}
		/** The first matching node anywhere below, depth-first. */
		function findFirst(node, name) {
			for (const child of elementChildren(node)) {
				if (localOf(child.name) === name) return child;
				const nested = findFirst(child, name);
				if (nested !== undefined) return nested;
			}
			return undefined;
		}
		/** All matching nodes below, depth-first. */
		function findAll(node, name, out = []) {
			for (const child of elementChildren(node)) {
				if (localOf(child.name) === name) out.push(child);
				findAll(child, name, out);
			}
			return out;
		}
		/** Concatenated text of a node's subtree. */
		function textOf(node) {
			if (typeof node === "string") return node;
			let text = "";
			for (const child of node.children) text += textOf(child);
			return text;
		}
		/** One attribute by its full (prefixed) name. */
		function attrOf(node, name) {
			const value = node.attrs[name];
			return value === undefined ? undefined : value;
		}
		//#endregion
		//#region src/client/zip.js
		/**
		 * A minimal ZIP reader: central-directory walk plus raw-DEFLATE inflate
		 * through the platform's own `DecompressionStream`.
		 *
		 * DOCX and XLSX are ZIP containers, so no Office library is needed — only
		 * the handful of XML parts each format keeps. The reader stays read-only and
		 * streams nothing: one entry is inflated on demand and released.
		 */
		const textDecoder = new TextDecoder("utf-8");
		/** Join byte chunks into one buffer. */
		function concatBytes(chunks, total) {
			const size = total === undefined ? chunks.reduce((sum, chunk) => sum + chunk.length, 0) : total;
			const joined = new Uint8Array(size);
			let offset = 0;
			for (const chunk of chunks) {
				joined.set(chunk, offset);
				offset += chunk.length;
			}
			return joined;
		}
		/** Inflate one raw-DEFLATE stream. */
		async function inflateRaw(bytes) {
			if (typeof DecompressionStream === "undefined") throw new Error("当前环境不支持 DecompressionStream，无法解压 Office 文件。");
			const stream = new DecompressionStream("deflate-raw");
			const writer = stream.writable.getWriter();
			const pumped = writer.write(bytes).then(() => writer.close());
			const reader = stream.readable.getReader();
			const chunks = [];
			for (;;) {
				const { done, value } = await reader.read();
				if (done) break;
				chunks.push(value);
			}
			await pumped;
			return concatBytes(chunks);
		}
		/**
		 * Index a ZIP container's central directory.
		 * @param bytes - the whole file.
		 * @returns readers for entry names, existence, bytes, and text.
		 */
		function readZip(bytes) {
			const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
			const minimum = Math.max(0, bytes.length - 66000);
			let end = -1;
			for (let at = bytes.length - 22; at >= minimum; at -= 1) {
				if (view.getUint32(at, true) === 0x06054b50) {
					end = at;
					break;
				}
			}
			if (end === -1) throw new Error("这不是一个有效的 OOXML 文件（未找到 ZIP 结尾记录）。");
			const count = view.getUint16(end + 10, true);
			const entries = /* @__PURE__ */ new Map();
			let at = view.getUint32(end + 16, true);
			for (let index = 0; index < count; index += 1) {
				if (at + 46 > bytes.length || view.getUint32(at, true) !== 0x02014b50) break;
				const method = view.getUint16(at + 10, true);
				const compressedSize = view.getUint32(at + 20, true);
				const size = view.getUint32(at + 24, true);
				const nameLength = view.getUint16(at + 28, true);
				const extraLength = view.getUint16(at + 30, true);
				const commentLength = view.getUint16(at + 32, true);
				const localOffset = view.getUint32(at + 42, true);
				const name = textDecoder.decode(bytes.subarray(at + 46, at + 46 + nameLength));
				entries.set(name, { method, compressedSize, size, localOffset });
				at += 46 + nameLength + extraLength + commentLength;
			}
			/** The regular files in container order. */
			const names = [...entries.keys()];
			/** Read one entry's uncompressed bytes. */
			const bytesOf = async (name) => {
				const entry = entries.get(name);
				if (entry === undefined) throw new Error(`Office 文件缺少内部条目 ${name}。`);
				if (view.getUint32(entry.localOffset, true) !== 0x04034b50) throw new Error(`Office 文件的内部条目 ${name} 头部无效。`);
				const nameLength = view.getUint16(entry.localOffset + 26, true);
				const extraLength = view.getUint16(entry.localOffset + 28, true);
				const start = entry.localOffset + 30 + nameLength + extraLength;
				const raw = bytes.subarray(start, start + entry.compressedSize);
				if (entry.method === 0) return raw;
				if (entry.method !== 8) throw new Error(`Office 文件的内部条目 ${name} 使用了不支持的压缩方式（${entry.method}）。`);
				return await inflateRaw(raw);
			};
			return {
				names,
				has: (name) => entries.has(name),
				bytes: bytesOf,
				text: async (name) => textDecoder.decode(await bytesOf(name)),
				find: (pattern) => names.find((name) => pattern.test(name))
			};
		}
		//#endregion
		//#region src/client/docx.js
		/**
		 * WordprocessingML to a render tree.
		 *
		 * Only what a reader needs survives the mapping: headings, paragraphs with
		 * inline emphasis, lists, indentation and tables. Images, fields, footnotes
		 * and revision marks collapse to their text, which keeps a duty roster or a
		 * notice readable without pretending to be a Word layout engine.
		 */
		/** Read `word/styles.xml` into a style-id → display-name table. */
		function readDocxStyles(xml) {
			const names = /* @__PURE__ */ new Map();
			if (xml === undefined) return names;
			for (const style of findAll(parseXml(xml), "style")) {
				const id = attrOf(style, "w:styleId");
				if (id === undefined) continue;
				const name = childNamed(style, "name");
				names.set(id, name === undefined ? id : attrOf(name, "w:val") ?? id);
			}
			return names;
		}
		/** The heading level a paragraph style announces, or zero. */
		function headingLevel(styleId, styleNames) {
			if (styleId === undefined) return 0;
			const candidates = [styleId, styleNames.get(styleId) ?? ""];
			for (const candidate of candidates) {
				const text = candidate.replace(/\s/gu, "");
				const match = /^(?:heading|标题)([1-6])$/iu.exec(text);
				if (match !== null) return Number(match[1]);
				if (/^[1-6]$/u.test(text)) return Number(text);
				if (/^head(?:ing)?[1-6]$/iu.test(text)) return Number(text.slice(-1));
			}
			return 0;
		}
		/** Inline runs of one paragraph, with the emphasis properties retained. */
		function readRuns(node) {
			const runs = [];
			const visit = (element) => {
				for (const child of elementChildren(element)) {
					const name = localOf(child.name);
					if (name === "r") {
						const properties = childNamed(child, "rPr");
						const emphasis = {
							bold: properties !== undefined && childNamed(properties, "b") !== undefined,
							italic: properties !== undefined && childNamed(properties, "i") !== undefined,
							underline: properties !== undefined && childNamed(properties, "u") !== undefined
						};
						let text = "";
						for (const piece of elementChildren(child)) {
							const pieceName = localOf(piece.name);
							if (pieceName === "t") text += textOf(piece);
							else if (pieceName === "tab") text += "\t";
							else if (pieceName === "br" || pieceName === "cr") text += "\n";
							else if (pieceName === "noBreakHyphen") text += "-";
						}
						if (text !== "") runs.push({ text, ...emphasis });
						continue;
					}
					if (name === "hyperlink" || name === "smartTag" || name === "ins" || name === "sdt" || name === "sdtContent") {
						visit(child);
						continue;
					}
					if (name === "del" || name === "drawing" || name === "pict" || name === "object") continue;
				}
			};
			visit(node);
			return runs;
		}
		/** Convert one `w:p` into a paragraph block. */
		function readParagraph(node, styleNames) {
			const properties = childNamed(node, "pPr");
			const styleId = properties === undefined ? undefined : attrOf(childNamed(properties, "pStyle") ?? { attrs: {} }, "w:val");
			const justification = properties === undefined ? undefined : attrOf(childNamed(properties, "jc") ?? { attrs: {} }, "w:val");
			const numbering = properties === undefined ? undefined : childNamed(properties, "numPr");
			const outline = properties === undefined ? undefined : attrOf(childNamed(properties, "outlineLvl") ?? { attrs: {} }, "w:val");
			const indent = properties === undefined ? undefined : attrOf(childNamed(properties, "ind") ?? { attrs: {} }, "w:left");
			let level = headingLevel(styleId, styleNames);
			if (level === 0 && outline !== undefined) level = Math.min(6, Number(outline) + 1);
			const listLevel = numbering === undefined ? -1 : Number(attrOf(childNamed(numbering, "ilvl") ?? { attrs: {} }, "w:val") ?? "0");
			return {
				kind: "p",
				level,
				align: justification === "center" ? "center" : justification === "right" ? "right" : justification === "both" ? "justify" : "start",
				indent: indent === undefined ? 0 : Math.round(Number(indent) / 15) + (listLevel > 0 ? listLevel * 18 : 0),
				bullet: listLevel < 0 ? "" : `${"　".repeat(listLevel)}• `,
				runs: readRuns(node)
			};
		}
		/** Convert one `w:tbl` into a table block, recursing into nested tables. */
		function readTable(node, styleNames) {
			const rows = [];
			for (const row of childrenNamed(node, "tr")) {
				const cells = [];
				for (const cell of childrenNamed(row, "tc")) {
					const blocks = [];
					for (const child of elementChildren(cell)) {
						const name = localOf(child.name);
						if (name === "p") blocks.push(readParagraph(child, styleNames));
						else if (name === "tbl") blocks.push(readTable(child, styleNames));
					}
					const span = attrOf(childNamed(cell, "tcPr") === undefined ? { attrs: {} } : childNamed(childNamed(cell, "tcPr"), "gridSpan") ?? { attrs: {} }, "w:val");
					cells.push({ blocks, span: span === undefined ? 1 : Math.max(1, Number(span)) });
				}
				rows.push(cells);
			}
			return { kind: "table", rows };
		}
		/** Convert the main document part into the render tree. */
		function readDocx(xml, stylesXml) {
			const styleNames = readDocxStyles(stylesXml);
			const root = parseXml(xml);
			const body = findFirst(root, "body");
			if (body === undefined) throw new Error("Word 文档缺少正文内容（word/document.xml 中没有 body）。");
			const blocks = [];
			for (const child of elementChildren(body)) {
				const name = localOf(child.name);
				if (name === "p") blocks.push(readParagraph(child, styleNames));
				else if (name === "tbl") blocks.push(readTable(child, styleNames));
			}
			return { blocks };
		}
		/** Parse a `.docx` container into the render tree. */
		async function parseDocx(bytes) {
			const zip = readZip(bytes);
			const main = zip.has("word/document.xml") ? "word/document.xml" : zip.find(/^word\/document\d*\.xml$/u);
			if (main === undefined) throw new Error("这不是一个有效的 Word 文档（缺少 word/document.xml）。");
			const styles = zip.has("word/styles.xml") ? await zip.text("word/styles.xml") : undefined;
			return readDocx(await zip.text(main), styles);
		}
		//#endregion
		//#region src/client/xlsx.js
		/**
		 * SpreadsheetML to a cell grid.
		 *
		 * The shared-string table, the style table's number formats, merges and
		 * column widths are the whole of what a preview needs: they decide whether a
		 * date reads as `2026-09-26` instead of `46251`, and whether a merged title
		 * row spans its columns. Formulas appear as their cached results, which is
		 * what a reader wants to see.
		 */
		/** Built-in number-format ids Excel reserves for dates and times. */
		const DATE_FORMATS = /* @__PURE__ */ new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);
		/** Whether a custom format code renders a date or a time. */
		function isDateFormatCode(code) {
			if (typeof code !== "string" || code === "") return false;
			const stripped = code.replace(/\[[^\]]*\]/gu, "").replace(/"[^"]*"/gu, "").replace(/\\./gu, "");
			return /[ymdhs]/iu.test(stripped);
		}
		/** Column index (0-based) of a cell reference such as `BC12`. */
		function columnIndex(ref) {
			let index = 0;
			let letters = 0;
			for (let at = 0; at < ref.length; at += 1) {
				const code = ref.charCodeAt(at);
				if (code >= 65 && code <= 90) index = index * 26 + (code - 64);
				else if (code >= 97 && code <= 122) index = index * 26 + (code - 96);
				else break;
				letters += 1;
			}
			return letters === 0 ? -1 : index - 1;
		}
		/** Column letters for a 0-based index. */
		function columnName(index) {
			let name = "";
			let value = index + 1;
			while (value > 0) {
				const remainder = (value - 1) % 26;
				name = String.fromCharCode(65 + remainder) + name;
				value = Math.floor((value - 1) / 26);
			}
			return name;
		}
		/** Render an Excel serial number as a date, or as a date and time. */
		function formatSerial(value, withTime) {
			if (!Number.isFinite(value)) return undefined;
			const days = Math.floor(value);
			const fraction = value - days;
			const base = Date.UTC(1899, 11, days < 60 ? 31 : 30) + days * 86400000;
			const date = new Date(base);
			const pad = (part) => String(part).padStart(2, "0");
			const day = `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}`;
			if (!withTime || fraction <= 0) return day;
			const seconds = Math.round(fraction * 86400);
			return `${day} ${pad(Math.floor(seconds / 3600) % 24)}:${pad(Math.floor(seconds / 60) % 60)}:${pad(seconds % 60)}`;
		}
		/** The style table's number formats, fonts and cell formats. */
		function readStyles(xml) {
			const formats = /* @__PURE__ */ new Map();
			const fonts = [];
			const cells = [];
			if (xml === undefined) return { formats, fonts, cells };
			const root = parseXml(xml);
			for (const format of findAll(root, "numFmt")) {
				const id = Number(attrOf(format, "numFmtId"));
				if (Number.isFinite(id)) formats.set(id, attrOf(format, "formatCode") ?? "");
			}
			for (const font of findAll(root, "font")) {
				fonts.push({
					bold: childNamed(font, "b") !== undefined,
					italic: childNamed(font, "i") !== undefined
				});
			}
			const cellXfs = findFirst(root, "cellXfs");
			for (const xf of cellXfs === undefined ? [] : elementChildren(cellXfs)) {
				const alignment = childNamed(xf, "alignment");
				cells.push({
					numFmtId: Number(attrOf(xf, "numFmtId") ?? "0"),
					fontId: Number(attrOf(xf, "fontId") ?? "0"),
					horizontal: alignment === undefined ? undefined : attrOf(alignment, "horizontal")
				});
			}
			return { formats, fonts, cells };
		}
		/** The shared-string table, in index order. */
		function readSharedStrings(xml) {
			if (xml === undefined) return [];
			return findAll(parseXml(xml), "si").map((item) => childrenNamed(item, "t").map((text) => textOf(text)).join("") || findAll(item, "t").map((text) => textOf(text)).join(""));
		}
		/** Resolve a sheet's part path from the workbook relationship it names. */
		function resolvePart(target, base) {
			if (target.startsWith("/")) return target.slice(1);
			const segments = `${base}/${target}`.split("/");
			const stack = [];
			for (const segment of segments) {
				if (segment === "" || segment === ".") continue;
				if (segment === "..") stack.pop();
				else stack.push(segment);
			}
			return stack.join("/");
		}
		/** The sheets a workbook declares, in tab order, with their part paths. */
		function readWorkbook(xml, relsXml) {
			const relations = /* @__PURE__ */ new Map();
			if (relsXml !== undefined) {
				for (const relation of findAll(parseXml(relsXml), "Relationship")) {
					const id = attrOf(relation, "Id");
					const target = attrOf(relation, "Target");
					if (id !== undefined && target !== undefined) relations.set(id, target);
				}
			}
			const sheets = [];
			for (const sheet of findAll(parseXml(xml), "sheet")) {
				const name = attrOf(sheet, "name") ?? `工作表${sheets.length + 1}`;
				const id = attrOf(sheet, "r:id") ?? attrOf(sheet, "id");
				const target = id === undefined ? undefined : relations.get(id);
				sheets.push({ name, part: target === undefined ? `xl/worksheets/sheet${sheets.length + 1}.xml` : resolvePart(target, "xl") });
			}
			return sheets;
		}
		/** Format one cell's stored value into display text. */
		function cellText(cell, shared, style) {
			const type = attrOf(cell, "t");
			const valueNode = childNamed(cell, "v");
			const raw = valueNode === undefined ? undefined : textOf(valueNode);
			if (type === "inlineStr") {
				const inline = childNamed(cell, "is");
				return inline === undefined ? "" : findAll(inline, "t").map((text) => textOf(text)).join("");
			}
			if (raw === undefined) return "";
			if (type === "s") {
				const entry = shared[Number(raw)];
				return entry === undefined ? raw : entry;
			}
			if (type === "b") return raw === "1" ? "TRUE" : "FALSE";
			if (type === "e") return raw;
			if (type === "str") return raw;
			const numeric = Number(raw);
			if (!Number.isFinite(numeric)) return raw;
			const isDate = style !== undefined && (DATE_FORMATS.has(style.numFmtId) || (style.formatCode !== undefined && isDateFormatCode(style.formatCode)));
			if (isDate) {
				const codeText = style?.formatCode ?? "";
				return formatSerial(numeric, /[hs]/iu.test(codeText) || DATE_FORMATS.has(style?.numFmtId ?? -1) && style.numFmtId >= 18) ?? raw;
			}
			return raw;
		}
		/** Parse one worksheet part into a sparse-then-dense row grid. */
		function readSheet(xml, shared, styles) {
			const rows = [];
			const merges = [];
			const root = parseXml(xml);
			const sheetData = findFirst(root, "sheetData");
			for (const row of sheetData === undefined ? [] : childrenNamed(sheetData, "row")) {
				const cells = [];
				for (const cell of childrenNamed(row, "c")) {
					const ref = attrOf(cell, "r") ?? "";
					const column = columnIndex(ref);
					if (column < 0) continue;
					const styleIndex = Number(attrOf(cell, "s") ?? "0");
					const definition = styles.cells[styleIndex];
					const font = definition === undefined ? undefined : styles.fonts[definition.fontId];
					const resolved = definition === undefined ? undefined : {
						numFmtId: definition.numFmtId,
						formatCode: styles.formats.get(definition.numFmtId),
						horizontal: definition.horizontal,
						bold: font?.bold === true,
						italic: font?.italic === true
					};
					const text = cellText(cell, shared, resolved);
					while (cells.length < column) cells.push(undefined);
					cells[column] = text === "" ? undefined : {
						text,
						bold: resolved?.bold === true,
						italic: resolved?.italic === true,
						align: resolved?.horizontal === "center" || resolved?.horizontal === "right" ? resolved.horizontal : undefined
					};
				}
				rows.push(cells);
			}
			for (const merge of findAll(root, "mergeCell")) {
				const ref = attrOf(merge, "ref");
				if (ref === undefined) continue;
				const [start, end] = ref.split(":");
				if (end === undefined) continue;
				const startColumn = columnIndex(start);
				const startRow = Number(/(\d+)$/u.exec(start)?.[1] ?? "0") - 1;
				const endColumn = columnIndex(end);
				const endRow = Number(/(\d+)$/u.exec(end)?.[1] ?? "0") - 1;
				if (startColumn < 0 || startRow < 0) continue;
				merges.push({ row: startRow, column: startColumn, rowSpan: endRow - startRow + 1, colSpan: endColumn - startColumn + 1 });
			}
			const widths = [];
			const cols = findFirst(root, "cols");
			for (const col of cols === undefined ? [] : childrenNamed(cols, "col")) {
				const min = Number(attrOf(col, "min") ?? "1");
				const max = Number(attrOf(col, "max") ?? String(min));
				const width = Number(attrOf(col, "width") ?? "0");
				for (let index = min; index <= Math.min(max, 512); index += 1) widths[index - 1] = width;
			}
			return { rows, merges, widths };
		}
		/** Parse an `.xlsx` container into sheets of display cells. */
		async function parseXlsx(bytes) {
			const zip = readZip(bytes);
			if (!zip.has("xl/workbook.xml")) throw new Error("这不是一个有效的 Excel 工作簿（缺少 xl/workbook.xml）。");
			const workbook = readWorkbook(await zip.text("xl/workbook.xml"), zip.has("xl/_rels/workbook.xml.rels") ? await zip.text("xl/_rels/workbook.xml.rels") : undefined);
			const shared = zip.has("xl/sharedStrings.xml") ? readSharedStrings(await zip.text("xl/sharedStrings.xml")) : [];
			const styles = readStyles(zip.has("xl/styles.xml") ? await zip.text("xl/styles.xml") : undefined);
			const sheets = [];
			for (const sheet of workbook) {
				let xml;
				try {
					xml = await zip.text(sheet.part);
				} catch {
					continue;
				}
				sheets.push({ name: sheet.name, ...readSheet(xml, shared, styles) });
			}
			if (sheets.length === 0) throw new Error("工作簿中没有任何可显示的工作表。");
			return { sheets };
		}
		//#endregion
		//#region src/client/limits.js
		/** Cells rendered per sheet before the grid is truncated. */
		const MAX_ROWS = 4000;
		const MAX_COLUMNS = 160;
		//#endregion
		//#region src/client/DocumentBody.js
		/** Render one paragraph block with its heading, list and alignment context. */
		function Paragraph({ block }) {
			const heading = block.level > 0;
			const style = {
				marginLeft: block.indent > 0 ? `${block.indent}px` : undefined,
				textAlign: block.align === "center" ? "center" : block.align === "right" ? "right" : undefined
			};
			const children = [];
			if (block.bullet !== "") children.push(block.bullet);
			block.runs.forEach((run, index) => {
				const emphasis = run.bold || run.italic || run.underline ? {
					fontWeight: run.bold ? 600 : undefined,
					fontStyle: run.italic ? "italic" : undefined,
					textDecoration: run.underline ? "underline" : undefined
				} : undefined;
				children.push(emphasis === undefined ? run.text : react.createElement("span", { key: index, style: emphasis }, run.text));
			});
			if (children.length === 0) children.push("");
			const className = `${heading ? "dso-h " : "dso-p "}${block.bullet === "" ? "" : "dso-list"}`.trim();
			const tag = heading ? `h${Math.min(6, block.level + 1)}` : "p";
			return react.createElement(tag, { className, style }, children);
		}
		/** Render a table block, recursing into nested tables. */
		function Table({ block }) {
			return react.createElement("table", { className: "dso-table" }, react.createElement("tbody", null, block.rows.map((row, rowIndex) => react.createElement("tr", { key: rowIndex }, row.map((cell, cellIndex) => react.createElement("td", {
				key: cellIndex,
				colSpan: cell.span > 1 ? cell.span : undefined
			}, cell.blocks.map((nested, index) => nested.kind === "table" ? react.createElement(Table, {
				key: index,
				block: nested
			}) : react.createElement(Paragraph, {
				key: index,
				block: nested
			}))))))));
		}
		/** The Word document pane: the parsed tree, or the reason it is not showing. */
		function WordBody({ blocks }) {
			return react.createElement("div", { className: "dso-body", "data-office-preview": "word" }, blocks.map((block, index) => block.kind === "table" ? react.createElement(Table, {
				key: index,
				block
			}) : react.createElement(Paragraph, {
				key: index,
				block
			})));
		}
		/** Render one worksheet as a bounded, merged-aware grid. */
		function SheetGrid({ sheet }) {
			const rows = sheet.rows.slice(0, MAX_ROWS);
			const covered = /* @__PURE__ */ new Set();
			const merges = /* @__PURE__ */ new Map();
			for (const merge of sheet.merges) merges.set(`${merge.row}:${merge.column}`, merge);
			let columns = 1;
			rows.forEach((row, rowIndex) => {
				const width = row.length + (merges.get(`${rowIndex}:${row.length}`)?.colSpan ?? 0);
				columns = Math.max(columns, Math.min(MAX_COLUMNS, Math.max(row.length, width)));
			});
			for (const merge of sheet.merges) {
				for (let row = merge.row; row < merge.row + merge.rowSpan; row += 1) {
					for (let column = merge.column; column < merge.column + merge.colSpan; column += 1) {
						if (row === merge.row && column === merge.column) continue;
						covered.add(`${row}:${column}`);
					}
				}
			}
			const header = react.createElement("thead", null, react.createElement("tr", null, react.createElement("th", { className: "dso-rowhead dso-corner" }, ""), Array.from({ length: columns }, (unused, index) => react.createElement("th", { key: index }, columnName(index)))));
			const body = [];
			rows.forEach((row, rowIndex) => {
				const cells = [react.createElement("td", { key: "head", className: "dso-rowhead" }, String(rowIndex + 1))];
				for (let column = 0; column < columns; column += 1) {
					const key = `${rowIndex}:${column}`;
					if (covered.has(key)) continue;
					const merge = merges.get(key);
					const cell = row[column];
					const classes = [cell?.bold ? "dso-b" : "", cell?.align === "center" ? "dso-c" : cell?.align === "right" ? "dso-r" : ""].filter(Boolean).join(" ");
					const width = sheet.widths[column];
					cells.push(react.createElement("td", {
						key: column,
						className: classes === "" ? undefined : classes,
						colSpan: merge !== undefined && merge.colSpan > 1 ? merge.colSpan : undefined,
						rowSpan: merge !== undefined && merge.rowSpan > 1 ? merge.rowSpan : undefined,
						style: width > 0 ? { minWidth: `${Math.round(Math.min(60, Math.max(6, width)) * 8)}px` } : undefined
					}, cell === undefined ? "" : cell.text));
				}
				body.push(react.createElement("tr", { key: rowIndex }, cells));
			});
			return react.createElement("div", { className: "dso-sheetwrap" }, react.createElement("table", { className: "dso-grid" }, header, react.createElement("tbody", null, body)), rows.length < sheet.rows.length ? react.createElement("p", { className: "dso-note" }, `仅显示前 ${MAX_ROWS} 行，共 ${sheet.rows.length} 行。`) : null);
		}
		/** The Excel workbook pane: one tab per sheet, the first one selected. */
		function WorkbookBody({ sheets, t }) {
			const [active, setActive] = react.useState(0);
			const index = Math.min(active, sheets.length - 1);
			const sheet = sheets[index];
			return react.createElement("div", { "data-office-preview": "excel" }, sheets.length > 1 ? react.createElement("div", { className: "dso-sheetbar", role: "tablist" }, sheets.map((entry, position) => react.createElement("button", {
				key: position,
				type: "button",
				role: "tab",
				className: "dso-sheet",
				"data-active": position === index,
				"aria-selected": position === index,
				onClick: () => setActive(position),
				title: entry.name
			}, entry.name))) : null, react.createElement(SheetGrid, { sheet, t }));
		}
		//#endregion
		//#region src/client/OfficeBody.js
		/** Human-readable size for the failure panel. */
		function formatSize(bytes) {
			if (bytes < 1024) return `${bytes} B`;
			if (bytes < 1048576) return `${Math.round(bytes / 1024)} KB`;
			return `${(bytes / 1048576).toFixed(1)} MB`;
		}
		/** The last path segment of an address, decoded when it can be. */
		function nameOf(address) {
			const segment = String(address).replaceAll("\\", "/");
			const name = segment.slice(segment.lastIndexOf("/") + 1);
			try {
				return decodeURIComponent(name);
			} catch {
				return name;
			}
		}
		/**
		 * The registered document body.
		 *
		 * The framework has already read the whole file — `loading: "bytes-complete"`
		 * means the pane hands over the bytes and nothing else — so this component
		 * only parses them, off the render path, and swaps in the result.
		 * @param props - document bytes, the tab address, the injected kind, locale.
		 * @returns the parsed document, a progress line, or the parse failure.
		 */
		function OfficeBody({ content, resourceAddress, kind, t }) {
			const data = content !== undefined && content.kind === "bytes" ? content.data : undefined;
			const name = react.useMemo(() => nameOf(resourceAddress), [resourceAddress]);
			const [state, setState] = react.useState({ status: "loading" });
			react.useEffect(() => {
				if (data === undefined) return undefined;
				let cancelled = false;
				setState({ status: "loading" });
				const parse = kind === "word" ? parseDocx(data) : parseXlsx(data);
				parse.then((parsed) => {
					if (!cancelled) setState({ status: "ready", parsed });
				}).catch((error) => {
					if (!cancelled) setState({
						status: "failed",
						message: error instanceof Error ? error.message : String(error)
					});
				});
				return () => {
					cancelled = true;
				};
			}, [data, kind]);
			if (data === undefined) return react.createElement("div", { className: "dso-status" }, react.createElement("p", null, t("empty")));
			if (state.status === "loading") return react.createElement("div", { className: "dso-status" }, react.createElement("p", null, t("loading")));
			if (state.status === "failed") return react.createElement("div", { className: "dso-status", role: "alert", "data-office-preview-error": true }, react.createElement("p", { className: "dso-status-name" }, name), react.createElement("p", null, t("failed")), react.createElement("p", { className: "dso-note" }, state.message), react.createElement("p", { className: "dso-note" }, `${formatSize(data.length)} · ${t("hint")}`));
			return react.createElement("div", { className: "dso-root", "data-office-preview-file": name }, kind === "word" ? react.createElement(WordBody, { blocks: state.parsed.blocks }) : react.createElement(WorkbookBody, { sheets: state.parsed.sheets, t }));
		}
		//#endregion
		//#region src/client/index.js
		/** This package's copy namespace. */
		const NS = "officeDocumentPreview";
		/** Simplified Chinese dictionary and key-set source of truth. */
		const zh = {
			"title.word": "Word 文档",
			"title.excel": "Excel 表格",
			loading: "正在解析 Office 文件…",
			empty: "文件内容不可用。",
			failed: "无法预览这个 Office 文件。",
			hint: "可在右侧“打开方式”中切换为纯文本。"
		};
		/** English dictionary, keyed identically. */
		const en = {
			"title.word": "Word document",
			"title.excel": "Excel workbook",
			loading: "Reading the Office file…",
			empty: "The file contents are unavailable.",
			failed: "This Office file could not be previewed.",
			hint: "Use “Open with” to switch to plain text."
		};
		/** Word implementation identity, shared by metadata and the keyed slot. */
		const WORD_ID = "dsh-office-preview/word";
		/** Excel implementation identity, shared by metadata and the keyed slot. */
		const EXCEL_ID = "dsh-office-preview/excel";
		/** WordprocessingML suffixes this renderer answers. */
		const WORD_EXTENSIONS = [
			"docx",
			"docm",
			"dotx",
			"dotm"
		];
		/** SpreadsheetML suffixes this renderer answers. */
		const EXCEL_EXTENSIONS = [
			"xlsx",
			"xlsm",
			"xltx",
			"xltm"
		];
		/**
		 * Describe one Office renderer independently from its keyed body slot.
		 * @param id - implementation identity.
		 * @param extensions - recognized suffixes.
		 * @param title - locale-owned implementation name.
		 * @returns metadata for a complete-bytes renderer.
		 */
		function definitionOf(id, extensions, title) {
			return {
				id,
				extensions,
				priority: "external",
				title,
				loading: "bytes-complete",
				wrap: false
			};
		}
		/**
		 * Required browser services: the preview registry, the slot registry, and copy.
		 */
		const inject = [
			"documentPreviews",
			"slots",
			"locale"
		];
		/**
		 * Client plugin body: register both renderers, their dictionaries, and the
		 * keyed document bodies that draw them.
		 * @param ctx - client root context carrying the preview and slot registries.
		 */
		function apply(ctx) {
			const t = ctx.locale.bind(NS);
			ctx.effect(() => ctx.locale.register(NS, {
				zh,
				en
			}), "office-preview: dictionaries");
			ctx.effect(() => ctx.documentPreviews.register(definitionOf(WORD_ID, WORD_EXTENSIONS, () => t("title.word"))), "office-preview: word metadata");
			ctx.effect(() => ctx.documentPreviews.register(definitionOf(EXCEL_ID, EXCEL_EXTENSIONS, () => t("title.excel"))), "office-preview: excel metadata");
			ctx.effect(() => ctx.slots.inject("sidebar.right.tab.document", () => ctx.slots.register({
				name: "sidebar.right.tab.document",
				key: WORD_ID,
				locale: NS,
				inject: () => ({ kind: "word" })
			}, OfficeBody)), "office-preview: word body");
			ctx.effect(() => ctx.slots.inject("sidebar.right.tab.document", () => ctx.slots.register({
				name: "sidebar.right.tab.document",
				key: EXCEL_ID,
				locale: NS,
				inject: () => ({ kind: "excel" })
			}, OfficeBody)), "office-preview: excel body");
		}
		exports.apply = apply;
		exports.inject = inject;
		exports.__test = {
			parseXml,
			textOf,
			readZip,
			parseDocx,
			parseXlsx,
			readDocx,
			readXlsxParts: { readStyles, readSharedStrings, readSheet, columnIndex, columnName, formatSerial, resolvePart, cellText },
			WordBody,
			WorkbookBody,
			SheetGrid,
			OfficeBody,
			MAX_ROWS,
			MAX_COLUMNS
		};
		return module.exports;
	}
});
