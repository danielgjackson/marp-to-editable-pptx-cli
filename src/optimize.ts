import { createHash } from 'node:crypto'
import path from 'node:path'
import JSZip from 'jszip'

/**
 * Post-processing of the PPTX that PptxGenJS writes, to make it smaller
 * without changing what it shows:
 *
 * 1. Identical media parts are stored once. PptxGenJS writes a separate copy
 *    of an image for every slide that uses it (e.g. the same background on
 *    every slide). Each media part is renamed after a hash of its content, so
 *    duplicates collapse to one part, and every relationship that pointed at
 *    a copy is redirected to it. Several relationships (from several slides)
 *    targeting one part is ordinary OOXML: PowerPoint does the same.
 * 2. The package is compressed. PptxGenJS stores every entry uncompressed;
 *    XML parts are deflated here. Already-compressed media (PNG, JPEG, ...)
 *    stays stored, as deflating it gains almost nothing.
 * 3. [Content_Types].xml loses <Override> entries for parts that don't exist
 *    (PptxGenJS lists a slide master per slide but writes only one).
 */

export interface OptimizeResult {
  data: Buffer
  mediaBefore: number
  mediaAfter: number
}

const MEDIA_DIR = 'ppt/media/'
const CONTENT_TYPES = '[Content_Types].xml'

/** Formats that are already compressed, so are stored rather than deflated. */
const PRECOMPRESSED = new Set([
  '.png', '.jpg', '.jpeg', '.gif', '.webp', '.mp4', '.m4v', '.mp3', '.m4a',
  '.xlsx', '.docx', '.pptx', '.zip',
])

/** The part a `.rels` file describes relationships for, e.g. ppt/slides/_rels/slide1.xml.rels → ppt/slides. */
function relsSourceDir(relsPath: string): string {
  // ".../_rels/<name>.rels" describes ".../<name>"; the root "_rels/.rels" describes the package.
  return path.posix.dirname(path.posix.dirname(relsPath)).replace(/^\.$/, '')
}

function safeDecode(s: string): string {
  try {
    return decodeURIComponent(s)
  } catch {
    return s
  }
}

/** Resolve a relationship target to a part name (no leading slash). */
function resolveTarget(sourceDir: string, target: string): string {
  const decoded = safeDecode(target)
  if (decoded.startsWith('/')) return path.posix.normalize(decoded).slice(1)
  return path.posix.normalize(path.posix.join(sourceDir || '.', decoded))
}

/** XML attribute value, from a single element's text. */
function attr(element: string, name: string): string | undefined {
  const m = new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(element)
  return m ? (m[1] ?? m[2]) : undefined
}

function setAttr(element: string, name: string, value: string): string {
  return element.replace(
    new RegExp(`(\\s${name}\\s*=\\s*)(?:"[^"]*"|'[^']*')`),
    (_, prefix: string) => `${prefix}"${value}"`,
  )
}

export async function optimizePptx(input: Buffer): Promise<OptimizeResult> {
  const zip = await JSZip.loadAsync(input)
  const entries = Object.values(zip.files).filter((f) => !f.dir)

  // Read every part once.
  const parts = new Map<string, { data: Buffer; date: Date }>()
  for (const entry of entries) {
    parts.set(entry.name, { data: await entry.async('nodebuffer'), date: entry.date })
  }

  // 1. Name each media part after its content: duplicates map to one name.
  const renamed = new Map<string, string>() // old part name → new part name
  const hashOfName = new Map<string, string>() // new part name → full hash
  let mediaBefore = 0
  for (const [name, { data }] of parts) {
    if (!name.startsWith(MEDIA_DIR) || name.slice(MEDIA_DIR.length).includes('/')) continue
    mediaBefore++
    const hash = createHash('sha256').update(data).digest('hex')
    const ext = path.posix.extname(name)
    // A short prefix of the hash keeps names readable; the full hash is
    // used only if two different contents would share a prefix.
    let newName = `${MEDIA_DIR}${hash.slice(0, 16)}${ext}`
    const taken = hashOfName.get(newName)
    if (taken !== undefined && taken !== hash) newName = `${MEDIA_DIR}${hash}${ext}`
    hashOfName.set(newName, hash)
    renamed.set(name, newName)
  }

  const out = new Map<string, { data: Buffer; date: Date }>()
  for (const [name, part] of parts) {
    const target = renamed.get(name) ?? name
    if (!out.has(target)) out.set(target, part)
  }

  // Redirect relationships that point at renamed media.
  for (const [name, part] of out) {
    if (!name.endsWith('.rels')) continue
    const sourceDir = relsSourceDir(name)
    const xml = part.data.toString('utf8')
    const updated = xml.replace(/<Relationship\b[^>]*>/g, (element) => {
      if (attr(element, 'TargetMode') === 'External') return element
      const target = attr(element, 'Target')
      if (target === undefined) return element
      const newName = renamed.get(resolveTarget(sourceDir, target))
      if (newName === undefined) return element
      const newTarget = target.startsWith('/')
        ? `/${newName}`
        : path.posix.relative(sourceDir, newName)
      return setAttr(element, 'Target', newTarget)
    })
    if (updated !== xml) out.set(name, { ...part, data: Buffer.from(updated, 'utf8') })
  }

  // 3. Content types: follow renames, and drop overrides for missing parts.
  const contentTypes = out.get(CONTENT_TYPES)
  if (contentTypes) {
    const present = new Set([...out.keys()].map((n) => n.toLowerCase()))
    const xml = contentTypes.data.toString('utf8')
    const updated = xml.replace(/<Override\b[^>]*\/>\s*/g, (element) => {
      const partName = attr(element, 'PartName')
      if (partName === undefined) return element
      const name = path.posix.normalize(safeDecode(partName)).replace(/^\//, '')
      const newName = renamed.get(name) ?? name
      if (!present.has(newName.toLowerCase())) return ''
      return newName === name ? element : setAttr(element, 'PartName', `/${newName}`)
    })
    out.set(CONTENT_TYPES, { ...contentTypes, data: Buffer.from(updated, 'utf8') })
  }

  // 2. Write a compressed package, [Content_Types].xml first as Office does.
  // Folder entries are left out: an OPC package doesn't need them.
  const result = new JSZip()
  const names = [...out.keys()].sort((a, b) =>
    a === CONTENT_TYPES ? -1 : b === CONTENT_TYPES ? 1 : 0,
  )
  for (const name of names) {
    const { data, date } = out.get(name)!
    const stored = PRECOMPRESSED.has(path.posix.extname(name).toLowerCase())
    result.file(name, data, {
      date,
      createFolders: false,
      compression: stored ? 'STORE' : 'DEFLATE',
      compressionOptions: stored ? null : { level: 9 },
    })
  }
  const data = await result.generateAsync({ type: 'nodebuffer' })
  return {
    data,
    mediaBefore,
    mediaAfter: new Set(renamed.values()).size,
  }
}
