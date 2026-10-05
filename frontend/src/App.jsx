import { useEffect, useMemo, useRef, useState } from 'react'
import { init as initPdfium } from '@embedpdf/pdfium'
import * as XLSX from 'xlsx'
import { PPTXViewer } from 'pptxjs'

const API = import.meta.env.VITE_API_URL || 'http://localhost:8000/api'

async function request(path, options = {}) {
  const response = await fetch(`${API}${path}`, options)
  const body = await response.json().catch(() => ({}))
  if (!response.ok) throw new Error(body.detail || '请求失败')
  return body
}

function Status({ value }) {
  const labels = { pending: '待处理', processing: '处理中', ready: '可检索', failed: '失败', not_supported: '暂不支持', not_indexed: '未建立索引' }
  return <span className={`status status-${value}`}>{labels[value] || value}</span>
}

function StorageStatus({ value }) {
  return value === 'missing' ? <span className="status status-missing">文件不存在</span> : null
}

function escapeRegExp(value) { return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') }

function formatBytes(bytes) {
  const value = Number(bytes)
  if (!Number.isFinite(value)) return '—'
  if (value < 1024 * 1024) return `${(value / 1024).toFixed(1)} KB`
  if (value < 1024 * 1024 * 1024) return `${(value / 1024 / 1024).toFixed(1)} MB`
  return `${(value / 1024 / 1024 / 1024).toFixed(2)} GB`
}

function Highlight({ text, query }) {
  if (!text || !query.trim()) return text
  const raw = query.trim()
  const compact = raw.replace(/\s+/g, '')
  const chinese = compact.match(/[\u4e00-\u9fff]/g)?.join('') || ''
  const terms = [raw, compact, ...(raw.match(/[A-Za-z0-9_]+/g) || [])]
  if (chinese.length > 2) for (let i = 0; i < chinese.length - 1; i += 1) terms.push(chinese.slice(i, i + 2))
  const unique = [...new Set(terms.filter(Boolean))].sort((a, b) => b.length - a.length)
  const pattern = new RegExp(`(${unique.map(escapeRegExp).join('|')})`, 'giu')
  return text.split(pattern).map((part, index) => unique.some(term => part.toLocaleLowerCase() === term.toLocaleLowerCase())
    ? <mark key={`${index}-${part}`}>{part}</mark> : <span key={`${index}-${part}`}>{part}</span>)
}

function resultSnippet(result) {
  const snippet = String(result?.snippets?.[0] || '').trim()
  const name = String(result?.document?.name || '').trim()
  if (!snippet || !name) return snippet
  if (snippet === name) return ''
  if (snippet.startsWith(name)) return snippet.slice(name.length).replace(/^[\s:：\-—]+/, '').trim()
  return snippet
}

let pdfiumPromise

function getPdfium() {
  if (!pdfiumPromise) {
    pdfiumPromise = fetch('/pdfium.wasm')
      .then(response => { if (!response.ok) throw new Error('PDFium 引擎文件读取失败'); return response.arrayBuffer() })
      .then(wasmBinary => initPdfium({ wasmBinary }))
      .then(instance => { instance.PDFiumExt_Init(); return instance })
  }
  return pdfiumPromise
}

function backgroundAssetUrl(version) {
  return `${API}/settings/background/file?v=${encodeURIComponent(version || Date.now())}`
}

function hexToRgba(value, alpha = 1) {
  const match = String(value || '').match(/^#([0-9a-f]{6})$/i)
  if (!match) return `rgba(23,59,69,${alpha})`
  const number = Number.parseInt(match[1], 16)
  return `rgba(${number >> 16},${(number >> 8) & 255},${number & 255},${alpha})`
}

async function renderPdfiumPages(bytes, maxPages = 20) {
  const pdfium = await getPdfium()
  const malloc = pdfium.pdfium.wasmExports.malloc
  const free = pdfium.pdfium.wasmExports.free
  const dataPtr = malloc(bytes.byteLength)
  // Emscripten may grow WebAssembly memory during malloc. Always refresh the
  // typed-array view after an allocation instead of reusing a detached view.
  pdfium.pdfium.HEAPU8.set(new Uint8Array(bytes), dataPtr)
  const document = pdfium.FPDF_LoadMemDocument(dataPtr, bytes.byteLength, '')
  if (!document) {
    free(dataPtr)
    throw new Error(`PDFium 无法打开 PDF（错误码 ${pdfium.FPDF_GetLastError()}）`)
  }
  const pageCount = pdfium.FPDF_GetPageCount(document)
  const pages = []
  try {
    for (let pageIndex = 0; pageIndex < Math.min(pageCount, maxPages); pageIndex += 1) {
      const page = pdfium.FPDF_LoadPage(document, pageIndex)
      if (!page) continue
      const sourceWidth = Math.max(1, pdfium.FPDF_GetPageWidthF(page))
      const sourceHeight = Math.max(1, pdfium.FPDF_GetPageHeightF(page))
      const scale = Math.min(1.5, 1100 / sourceWidth)
      const width = Math.max(1, Math.ceil(sourceWidth * scale))
      const height = Math.max(1, Math.ceil(sourceHeight * scale))
      const bitmap = pdfium.FPDFBitmap_CreateEx(width, height, 4, 0, 0)
      if (!bitmap) {
        pdfium.FPDF_ClosePage(page)
        continue
      }
      pdfium.FPDFBitmap_FillRect(bitmap, 0, 0, width, height, 0xffffffff)
      pdfium.FPDF_RenderPageBitmap(bitmap, page, 0, 0, width, height, 0, 0)
      const bufferPtr = pdfium.FPDFBitmap_GetBuffer(bitmap)
      const stride = pdfium.FPDFBitmap_GetStride(bitmap)
      // Rendering can grow the WASM heap as well; slice the current view so
      // the source bytes remain detached from PDFium's mutable memory.
      const source = pdfium.pdfium.HEAPU8.slice(bufferPtr, bufferPtr + stride * height)
      const pixels = new Uint8ClampedArray(width * height * 4)
      for (let y = 0; y < height; y += 1) {
        for (let x = 0; x < width; x += 1) {
          const sourceOffset = y * stride + x * 4
          const targetOffset = (y * width + x) * 4
          pixels[targetOffset] = source[sourceOffset + 2]
          pixels[targetOffset + 1] = source[sourceOffset + 1]
          pixels[targetOffset + 2] = source[sourceOffset]
          pixels[targetOffset + 3] = source[sourceOffset + 3]
        }
      }
      pages.push({ index: pageIndex + 1, width, height, pixels })
      pdfium.FPDFBitmap_Destroy(bitmap)
      pdfium.FPDF_ClosePage(page)
    }
    return { pages, pageCount }
  } finally {
    pdfium.FPDF_CloseDocument(document)
    free(dataPtr)
  }
}

function PdfiumPage({ page }) {
  const canvasRef = useMemo(() => ({ current: null }), [])
  useEffect(() => {
    const canvas = canvasRef.current
    if (!canvas) return
    canvas.width = page.width
    canvas.height = page.height
    const context = canvas.getContext('2d', { alpha: false })
    context.putImageData(new ImageData(page.pixels, page.width, page.height), 0, 0)
  }, [canvasRef, page])
  return <canvas ref={node => { canvasRef.current = node }} className="pdfium-page" aria-label={`第 ${page.index} 页`} />
}

function SpreadsheetPreview({ src, fallbackHtml }) {
  const [sheets, setSheets] = useState(null)
  const [loadError, setLoadError] = useState('')
  useEffect(() => {
    let active = true
    setSheets(null); setLoadError('')
    fetch(src)
      .then(response => { if (!response.ok) throw new Error('表格原文读取失败'); return response.arrayBuffer() })
      .then(buffer => {
        const workbook = XLSX.read(buffer, { type: 'array', cellDates: true })
        return workbook.SheetNames.map(name => ({ name, rows: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '' }).slice(0, 5000) }))
      })
      .then(value => { if (active) setSheets(value) })
      .catch(error => { if (active) setLoadError(error.message || '表格原文读取失败') })
    return () => { active = false }
  }, [src])
  if (loadError) return fallbackHtml ? <div className="rich-preview" dangerouslySetInnerHTML={{ __html: fallbackHtml }} /> : <div className="empty">{loadError}，可切换到纯文本或使用下载按钮。</div>
  if (!sheets) return <div className="preview-loading">正在使用 SheetJS 读取表格原文…</div>
  return <div className="rich-preview office-preview">{sheets.map(sheet => <section key={sheet.name}><h3>{sheet.name}</h3>{sheet.rows.length ? <table><tbody>{sheet.rows.map((row, rowIndex) => <tr key={rowIndex}>{row.map((cell, cellIndex) => rowIndex === 0 ? <th key={cellIndex}>{String(cell)}</th> : <td key={cellIndex}>{String(cell)}</td>)}</tr>)}</tbody></table> : <p>此工作表没有可显示内容。</p>}</section>)}</div>
}

function PptxPreview({ src, fallbackHtml }) {
  const containerRef = useRef(null)
  const [loadError, setLoadError] = useState('')
  const [loaded, setLoaded] = useState(false)
  useEffect(() => {
    let viewer
    let active = true
    setLoadError(''); setLoaded(false)
    if (!containerRef.current) return undefined
    try {
      viewer = new PPTXViewer(containerRef.current, { showControls: true, keyboardNavigation: true, onLoad: () => { if (active) setLoaded(true) }, onError: error => { if (active) setLoadError(error?.message || 'PPTX 原文读取失败') } })
      viewer.load(src).then(() => { if (active) setLoaded(true) }).catch(error => { if (active) setLoadError(error?.message || 'PPTX 原文读取失败') })
    } catch (error) {
      setLoadError(error?.message || 'PPTX 预览初始化失败')
    }
    return () => { active = false; viewer?.destroy?.() }
  }, [src])
  if (loadError && fallbackHtml) return <div className="rich-preview" dangerouslySetInnerHTML={{ __html: fallbackHtml }} />
  if (loadError) return <div className="empty">{loadError}，可切换到纯文本或使用下载按钮。</div>
  return <div className="pptx-preview" ref={containerRef}>{!loaded && <div className="preview-loading">正在使用 PptxJS 读取幻灯片原文…</div>}</div>
}

function PdfPreview({ src }) {
  const [rendered, setRendered] = useState(null)
  const [loadError, setLoadError] = useState('')
  useEffect(() => {
    let active = true
    setRendered(null); setLoadError('')
    fetch(src)
      .then(response => { if (!response.ok) throw new Error('PDF 原文读取失败'); return response.json() })
      .then(payload => {
        if (!payload?.data) throw new Error('PDF 原文数据为空')
        const binary = atob(payload.data)
        const bytes = new Uint8Array(binary.length)
        for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index)
        return bytes.buffer
      })
      .then(bytes => renderPdfiumPages(bytes))
      .then(result => { if (active) setRendered(result) })
      .catch(error => { if (active) setLoadError(error.message || 'PDFium 原文渲染失败') })
    return () => { active = false }
  }, [src])
  if (loadError) return <div className="empty">{loadError}，可切换到纯文本或使用下载按钮。</div>
  if (!rendered) return <div className="preview-loading">正在使用 PDFium 引擎渲染 PDF 原文…</div>
  return <div className="pdfium-preview"><div className="pdf-preview-note"><span>共 {rendered.pageCount} 页{rendered.pageCount > rendered.pages.length ? `，当前显示前 ${rendered.pages.length} 页` : ''}</span></div><div className="pdfium-pages">{rendered.pages.map(page => <PdfiumPage key={page.index} page={page} />)}</div></div>
}

export default function App() {
  const [categories, setCategories] = useState([])
  const [documents, setDocuments] = useState([])
  const [archived, setArchived] = useState(false)
  const [category, setCategory] = useState('')
  const [query, setQuery] = useState('')
  const [mode, setMode] = useState('filename')
  const [results, setResults] = useState(null)
  const [message, setMessage] = useState('')
  const [error, setError] = useState('')
  const [uploading, setUploading] = useState(false)
  const [scanningStorage, setScanningStorage] = useState(false)
  const [selectedIds, setSelectedIds] = useState([])
  const [preview, setPreview] = useState(null)
  const [previewMode, setPreviewMode] = useState('rich')
  const [databaseOverview, setDatabaseOverview] = useState(null)
  const [storageInfo, setStorageInfo] = useState(null)
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [settingsDraft, setSettingsDraft] = useState(null)
  const [backgroundUrl, setBackgroundUrl] = useState('')
  const [backgroundUploading, setBackgroundUploading] = useState(false)
  const [savingSettings, setSavingSettings] = useState(false)
  const [stoppingServices, setStoppingServices] = useState(false)
  const [newCategoryName, setNewCategoryName] = useState('')
  const [newCategoryParentId, setNewCategoryParentId] = useState('')
  const [editingCategoryId, setEditingCategoryId] = useState(null)
  const [editingCategoryName, setEditingCategoryName] = useState('')
  const [editingCategoryParentId, setEditingCategoryParentId] = useState('')
  const [categoryBusy, setCategoryBusy] = useState(false)
  const [categoryOpen, setCategoryOpen] = useState(false)
  const [closingModal, setClosingModal] = useState('')
  const [sortMethod, setSortMethod] = useState('modified_desc')
  const [pageSize, setPageSize] = useState(20)
  const [semanticResultLimit, setSemanticResultLimit] = useState(5)
  const [showAllResults, setShowAllResults] = useState(false)
  const [currentPage, setCurrentPage] = useState(1)
  const [formatFilters, setFormatFilters] = useState([])
  const [dateFrom, setDateFrom] = useState('')
  const [dateTo, setDateTo] = useState('')
  const [batchCategoryId, setBatchCategoryId] = useState('')
  const [batchUpdating, setBatchUpdating] = useState(false)
  const [renameOpen, setRenameOpen] = useState(false)
  const [renameDraft, setRenameDraft] = useState({ find_text: '', replace_text: '', prefix: '', suffix: '', extension: '' })
  const [renaming, setRenaming] = useState(false)
  const [inlineRenameId, setInlineRenameId] = useState(null)
  const [inlineRenameValue, setInlineRenameValue] = useState('')
  const [inlineRenameBusy, setInlineRenameBusy] = useState(false)
  const [lastRenameUndo, setLastRenameUndo] = useState(null)
  // The list/tree switch is purely a presentation choice; both views use the
  // same persisted document records and update endpoint.
  const [libraryView, setLibraryView] = useState('list')
  const [expandedCategoryIds, setExpandedCategoryIds] = useState(() => new Set(['__uncategorized']))

  const orderedCategories = useMemo(() => {
    const children = new Map()
    categories.forEach(item => {
      const parent = item.parent_id == null ? '__root__' : String(item.parent_id)
      if (!children.has(parent)) children.set(parent, [])
      children.get(parent).push(item)
    })
    for (const values of children.values()) values.sort((a, b) => a.name.localeCompare(b.name, 'zh-Hans'))
    const result = []
    const visited = new Set()
    const walk = (parent, depth) => {
      for (const item of children.get(parent) || []) {
        if (visited.has(item.id)) continue
        visited.add(item.id)
        result.push({ item, depth })
        walk(String(item.id), depth + 1)
      }
    }
    walk('__root__', 0)
    // Keep malformed or legacy records visible even if their parent is missing.
    categories.forEach(item => { if (!visited.has(item.id)) result.push({ item, depth: 0 }) })
    return result
  }, [categories])
  const categoryPath = (categoryId) => {
    if (categoryId == null || categoryId === '') return '未分类'
    const byId = new Map(categories.map(item => [String(item.id), item]))
    const parts = []
    const seen = new Set()
    let current = byId.get(String(categoryId))
    while (current && !seen.has(current.id)) {
      seen.add(current.id); parts.unshift(current.name); current = current.parent_id == null ? null : byId.get(String(current.parent_id))
    }
    return parts.length ? parts.join(' / ') : '未分类'
  }

  const documentQuery = () => {
    const suffix = new URLSearchParams({ archived: String(archived) })
    if (category) suffix.set('category_id', category)
    return `/documents?${suffix}`
  }
  const refreshDocuments = async (preserveSelection = true) => {
    const docs = await request(documentQuery())
    setDocuments(docs)
    if (preserveSelection) setSelectedIds(ids => ids.filter(id => docs.some(doc => doc.id === id)))
    else setSelectedIds([])
  }
  const refresh = async () => {
    const [cats, storage] = await Promise.all([request('/categories'), refreshDocuments(false), request('/storage/status')])
    setCategories(cats)
    setStorageInfo(storage)
  }
  useEffect(() => { refresh().catch(e => setError(e.message)) }, [archived, category])
  const [themeSettings, setThemeSettings] = useState({ workspace_opacity: 0.96, background_opacity: 0.38, accent_color: '#173b45', accent_opacity: 1, modal_animation: 'fade', eyebrow_text: 'Documents Workspace', eyebrow_color: '#9ad4c8', title_text: '文件管理与知识检索平台', subtitle_text: 'Made By KOKONA', title_color: '#f5fbfa', subtitle_color: '#9fc4c9' })
  useEffect(() => { request('/settings').then(saved => { setSortMethod(saved.sort_method || 'modified_desc'); setPageSize(Number.isFinite(Number(saved.page_size)) ? Number(saved.page_size) : 20); setSemanticResultLimit(Number.isFinite(Number(saved.semantic_result_limit)) ? Number(saved.semantic_result_limit) : 5); setThemeSettings({ workspace_opacity: Number(saved.workspace_opacity ?? 0.96), background_opacity: Number(saved.background_opacity ?? 0.38), accent_color: saved.accent_color || '#173b45', accent_opacity: Number(saved.accent_opacity ?? 1), modal_animation: saved.modal_animation || 'fade', eyebrow_text: saved.eyebrow_text ?? 'Documents Workspace', eyebrow_color: saved.eyebrow_color || '#9ad4c8', title_text: saved.title_text ?? '文件管理与知识检索平台', subtitle_text: saved.subtitle_text ?? 'Made By KOKONA', title_color: saved.title_color || '#f5fbfa', subtitle_color: saved.subtitle_color || '#9fc4c9' }) }).catch(() => {}) }, [])
  useEffect(() => {
    request('/settings/background').then(saved => setBackgroundUrl(saved.enabled ? backgroundAssetUrl(saved.version) : '')).catch(() => {})
  }, [])
  const hasIndexing = documents.some(doc => doc.index_status === 'pending' || doc.index_status === 'processing')
  useEffect(() => {
    if (!hasIndexing) return undefined
    const timer = window.setInterval(() => refreshDocuments(true).catch(() => {}), 1600)
    return () => window.clearInterval(timer)
  }, [archived, category, hasIndexing])
  // Files can be removed directly from the storage folder.  Refresh the list
  // periodically so the persisted status becomes visible without requiring a
  // browser refresh or a manual scan click.
  useEffect(() => {
    const timer = window.setInterval(() => refreshDocuments(true).catch(() => {}), 5000)
    return () => window.clearInterval(timer)
  }, [archived, category])
  const migrationStatus = storageInfo?.migration?.status || 'idle'
  useEffect(() => {
    if (!['queued', 'processing'].includes(migrationStatus)) return undefined
    const timer = window.setInterval(async () => {
      try {
        const status = await request('/storage/migration-status')
        setStorageInfo(previous => previous ? { ...previous, migration: status } : previous)
        if (status.status === 'completed') {
          await refresh()
          setMessage('数据库和文档已移动，服务已重新启动。')
        } else if (status.status === 'failed') {
          setError(status.message || '存储迁移失败，请查看 runtime/.storage-migration.log')
        }
      } catch (_) {
        // The API is expected to restart during migration; the next poll
        // observes the final status after it comes back.
      }
    }, 1800)
    return () => window.clearInterval(timer)
  }, [migrationStatus])

  const createCategory = async () => {
    const name = newCategoryName.trim()
    if (!name) { setError('请输入分类名称'); return }
    try {
      setCategoryBusy(true); setError('')
      await request('/categories', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name, parent_id: newCategoryParentId ? Number(newCategoryParentId) : null }) })
      setNewCategoryName(''); setNewCategoryParentId(''); setMessage(`已添加分类“${name}”`); await refresh()
    } catch (e) { setError(e.message) } finally { setCategoryBusy(false) }
  }
  const beginRenameCategory = (item) => { setEditingCategoryId(item.id); setEditingCategoryName(item.name); setEditingCategoryParentId(item.parent_id ? String(item.parent_id) : ''); setError('') }
  const cancelRenameCategory = () => { setEditingCategoryId(null); setEditingCategoryName(''); setEditingCategoryParentId('') }
  const saveRenameCategory = async (item) => {
    const name = editingCategoryName.trim()
    if (!name) { setError('分类名称不能为空'); return }
    if (name === item.name) { cancelRenameCategory(); return }
    try {
      setCategoryBusy(true); setError('')
      await request(`/categories/${item.id}`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name, parent_id: editingCategoryParentId ? Number(editingCategoryParentId) : null }) })
      setMessage(`分类已修改为“${name}”`); cancelRenameCategory(); await refresh()
    } catch (e) { setError(e.message) } finally { setCategoryBusy(false) }
  }
  const deleteCategory = async (target) => {
    const visibleCount = documents.filter(doc => doc.category_id === target.id).length
    if (!window.confirm(`确定删除分类“${target.name}”吗？${visibleCount} 个当前列表文件会回到未分类，原文件不会删除。`)) return
    try {
      setCategoryBusy(true); setError('')
      const result = await request(`/categories/${target.id}`, { method: 'DELETE' })
      setMessage(`分类已删除，${result.documents_moved_to_default} 个文件已移动到未分类`)
      if (String(target.id) === category) setCategory('')
      await refresh()
    } catch (e) { setError(e.message) } finally { setCategoryBusy(false) }
  }

  const upload = async (event) => {
    const files = [...(event.target.files || [])]
    if (!files.length) return
    setUploading(true); setError(''); setMessage(`正在上传 ${files.length} 个文件…`)
    const sendUpload = async (selectedFiles, duplicateAction = 'prompt') => {
      const form = new FormData(); selectedFiles.forEach(file => form.append('files', file)); form.append('duplicate_action', duplicateAction)
      return request('/documents/batch', { method: 'POST', body: form })
    }
    try {
      let result = await sendUpload(files)
      let uploaded = [...(result.uploaded || [])]
      let failedItems = [...(result.failed || [])]
      let duplicateCount = result.duplicates?.length || 0
      if (duplicateCount) {
        const duplicateLines = result.duplicates.map(item => `“${item.name}”与已有文件“${item.existing_name}”内容完全相同`).join('\n')
        const keepRenamed = window.confirm(`检测到 ${duplicateCount} 个内容重复文件：\n${duplicateLines}\n\n点击“确定”将以“（副本）”重命名后保留；点击“取消”只保留已有文件。`)
        if (keepRenamed) {
          const duplicateIndexes = new Set(result.duplicates.map(item => item.index))
          const duplicateFiles = files.filter((_, index) => duplicateIndexes.has(index))
          const renamed = await sendUpload(duplicateFiles, 'rename')
          uploaded = [...uploaded, ...(renamed.uploaded || [])]
          failedItems = [...failedItems, ...(renamed.failed || [])]
          duplicateCount = renamed.duplicates?.length || 0
        }
      }
      const failed = failedItems.length
      setMessage(`已保存 ${uploaded.length} 个文件${duplicateCount ? `，跳过 ${duplicateCount} 个重复文件` : ''}${failed ? `，${failed} 个失败` : ''}；索引完成后会自动显示“可检索”`)
      if (failed) setError(failedItems.map(item => `${item.name}：${item.error}`).join('；'))
      // The upload response already contains the persisted document rows. Add
      // them to the current view immediately; the follow-up refresh reconciles
      // categories/storage metadata without making the user wait for it.
      const visibleUploads = uploaded.filter(item => (
        Boolean(item.archived) === archived && (!category || String(item.category_id || '') === String(category))
      ))
      if (visibleUploads.length) {
        setDocuments(previous => {
          const byId = new Map(previous.map(item => [item.id, item]))
          visibleUploads.forEach(item => byId.set(item.id, item))
          return [...byId.values()]
        })
      }
      refresh().catch(e => setError(e.message))
    } catch (e) { setError(e.message) } finally { setUploading(false); event.target.value = '' }
  }

  const search = async (event) => {
    event?.preventDefault(); if (!query.trim()) return
    try { setResults(await request(`/search?${new URLSearchParams({ q: query, mode, ...(category ? { category_id: category } : {}), archived: String(archived) })}`)); setShowAllResults(false); setError('') } catch (e) { setError(e.message) }
  }

  const update = async (doc, changes, silent = false) => {
    try {
      const updated = await request(`/documents/${doc.id}`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(changes) })
      setMessage(updated.storage_conflict ? '更改文件目录时发现目标目录存在同名文件，系统已自动保留两个文件。' : '')
      await refresh()
    } catch (e) { setError(e.message) }
  }
  const scanStorage = async () => {
    if (scanningStorage) return
    setScanningStorage(true); setError('')
    try {
      const result = await request('/storage/scan', { method: 'POST' })
      setMessage(result.missing_count ? `扫描完成：发现 ${result.missing_count} 个文件不存在，请在列表中删除对应索引。` : `扫描完成：已检查 ${result.checked} 个文件，文件均存在。`)
      await refresh()
    } catch (e) { setError(e.message) } finally { setScanningStorage(false) }
  }
  const deleteMissingDocument = async (doc) => {
    if (doc.storage_status !== 'missing') return
    try {
      // The source file is already gone. Remove its database record and the
      // related chunks/jobs together, without asking for a second confirmation.
      await request('/documents/batch', { method: 'DELETE', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ids: [doc.id] }) })
      setMessage(`已删除“${doc.name}”的文件记录，并一并删除索引。`)
      await refresh()
    } catch (e) { setError(e.message) }
  }
  const retry = async (doc) => { try { await request(`/documents/${doc.id}/index-retry`, { method: 'POST' }); await refresh() } catch (e) { setError(e.message) } }
  const openPreview = async (doc) => { try { setError(''); setPreviewMode('rich'); setPreview(await request(`/documents/${doc.id}/preview`)) } catch (e) { setError(e.message) } }
  const openDatabase = async () => { try { setError(''); setDatabaseOverview(await request('/database/overview')) } catch (e) { setError(e.message) } }
  const closeFloating = (name, action) => {
    if (closingModal) return
    setClosingModal(name)
    window.setTimeout(() => { action(); setClosingModal('') }, 280)
  }
  const closeSettings = () => closeFloating('settings', () => setSettingsOpen(false))
  const closeCategory = () => closeFloating('category', () => setCategoryOpen(false))
  const closePreview = () => closeFloating('preview', () => setPreview(null))
  const closeDatabase = () => closeFloating('database', () => setDatabaseOverview(null))
  const closeRename = () => closeFloating('rename', () => setRenameOpen(false))
  const openSettings = async () => {
    try {
      setError('')
      const [saved, storage] = await Promise.all([request('/settings'), request('/storage/status')])
      setStorageInfo(storage)
      setThemeSettings({ workspace_opacity: Number(saved.workspace_opacity ?? 0.96), background_opacity: Number(saved.background_opacity ?? 0.38), accent_color: saved.accent_color || '#173b45', accent_opacity: Number(saved.accent_opacity ?? 1), modal_animation: saved.modal_animation || 'fade', eyebrow_text: saved.eyebrow_text ?? 'Documents Workspace', eyebrow_color: saved.eyebrow_color || '#9ad4c8', title_text: saved.title_text ?? '文件管理与知识检索平台', subtitle_text: saved.subtitle_text ?? 'Made By KOKONA', title_color: saved.title_color || '#f5fbfa', subtitle_color: saved.subtitle_color || '#9fc4c9' })
      setSettingsDraft({ ...saved, database_host_path: storage.database_host_path || './runtime/postgres', storage_host_path: storage.host_path || './runtime/storage', page_size: Number.isFinite(Number(saved.page_size)) ? Number(saved.page_size) : 20, semantic_result_limit: Number.isFinite(Number(saved.semantic_result_limit)) ? Number(saved.semantic_result_limit) : 5, max_file_mb: (saved.max_file_bytes / 1024 / 1024).toFixed(0), max_batch_download_gb: (saved.max_batch_download_bytes / 1024 / 1024 / 1024).toFixed(1), index_poll_seconds: String(saved.index_poll_seconds), embedding_api_key: '' })
      setSettingsOpen(true)
    } catch (e) { setError(e.message) }
  }
  const uploadBackground = async (event) => {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) return
    setBackgroundUploading(true); setError('')
    const form = new FormData(); form.append('file', file)
    try {
      const saved = await request('/settings/background', { method: 'POST', body: form })
      setBackgroundUrl(backgroundAssetUrl(saved.version))
      setMessage('操作台背景已更新并持久化保存。')
    } catch (e) { setError(e.message) } finally { setBackgroundUploading(false) }
  }
  const clearBackground = async () => {
    try {
      setError('')
      await request('/settings/background', { method: 'DELETE' })
      setBackgroundUrl('')
      setMessage('已清除操作台背景，恢复默认背景。')
    } catch (e) { setError(e.message) }
  }
  const saveSettings = async () => {
    if (!settingsDraft) return
    const currentDatabasePath = storageInfo?.database_host_path || './runtime/postgres'
    const currentStoragePath = storageInfo?.host_path || './runtime/storage'
    const nextDatabasePath = String(settingsDraft.database_host_path || '').trim()
    const nextStoragePath = String(settingsDraft.storage_host_path || '').trim()
    const pathChanged = nextDatabasePath !== currentDatabasePath || nextStoragePath !== currentStoragePath
    if (pathChanged && !window.confirm(`确认立即移动数据库和文档吗？平台会暂时停止 Docker 服务，将数据库移动到“${nextDatabasePath}”，文档移动到“${nextStoragePath}”，完成后自动重新启动。目标目录必须为空。`)) return
    try {
      setSavingSettings(true); setError('')
      const payload = { ...settingsDraft, max_file_bytes: Math.round(Number(settingsDraft.max_file_mb) * 1024 * 1024), max_batch_download_bytes: Math.round(Number(settingsDraft.max_batch_download_gb) * 1024 * 1024 * 1024), index_poll_seconds: Number(settingsDraft.index_poll_seconds), page_size: Number(settingsDraft.page_size), sort_method: settingsDraft.sort_method || 'modified_desc' }
      delete payload.max_file_mb; delete payload.max_batch_download_gb; delete payload.embedding_api_key_set; delete payload.database_host_path; delete payload.storage_host_path
      if (!payload.embedding_api_key) delete payload.embedding_api_key
      const saved = await request('/settings', { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) })
      let migration = null
      if (pathChanged) {
        migration = await request('/storage/migrate', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ database_host_path: nextDatabasePath, storage_host_path: nextStoragePath }) })
        setStorageInfo(previous => previous ? { ...previous, migration } : previous)
      }
      setSettingsDraft({ ...saved, database_host_path: nextDatabasePath || currentDatabasePath, storage_host_path: nextStoragePath || currentStoragePath, page_size: Number.isFinite(Number(saved.page_size)) ? Number(saved.page_size) : pageSize, semantic_result_limit: Number.isFinite(Number(saved.semantic_result_limit)) ? Number(saved.semantic_result_limit) : semanticResultLimit, max_file_mb: (saved.max_file_bytes / 1024 / 1024).toFixed(0), max_batch_download_gb: (saved.max_batch_download_bytes / 1024 / 1024 / 1024).toFixed(1), index_poll_seconds: String(saved.index_poll_seconds), embedding_api_key: '' })
      setSortMethod(saved.sort_method || 'modified_desc')
      setPageSize(Number.isFinite(Number(saved.page_size)) ? Number(saved.page_size) : pageSize)
      setSemanticResultLimit(Number.isFinite(Number(saved.semantic_result_limit)) ? Number(saved.semantic_result_limit) : semanticResultLimit)
      setThemeSettings({ workspace_opacity: Number(saved.workspace_opacity ?? 0.96), background_opacity: Number(saved.background_opacity ?? 0.38), accent_color: saved.accent_color || '#173b45', accent_opacity: Number(saved.accent_opacity ?? 1), modal_animation: saved.modal_animation || 'fade', eyebrow_text: saved.eyebrow_text ?? 'Documents Workspace', eyebrow_color: saved.eyebrow_color || '#9ad4c8', title_text: saved.title_text ?? '文件管理与知识检索平台', subtitle_text: saved.subtitle_text ?? 'Made By KOKONA', title_color: saved.title_color || '#f5fbfa', subtitle_color: saved.subtitle_color || '#9fc4c9' })
      if (migration) setMessage('迁移请求已确认，正在停止服务并移动数据库和文档；完成后会自动重新启动。')
      else { setStorageInfo(await request('/storage/status')); setMessage('设置已保存，后续上传、分类和索引任务立即使用新设置。') }
      closeSettings()
    } catch (e) { setError(e.message) } finally { setSavingSettings(false) }
  }
  const batchCategorize = async () => {
    if (!selectedIds.length) return
    const categoryId = batchCategoryId ? Number(batchCategoryId) : null
    const action = storageInfo?.layout === 'folders' ? '源文件会同步移动到新的分类目录，原文件不会删除。' : '只会更新数据库分类，源文件保留在统一持久化目录，不会删除。'
    if (!window.confirm(`确定将选中的 ${selectedIds.length} 个文件归入${categoryId ? '所选分类' : '未分类'}吗？${action}`)) return
    try {
      setBatchUpdating(true); setError(''); setMessage(`正在批量分类 ${selectedIds.length} 个文件…`)
      await Promise.all(selectedIds.map(id => request(`/documents/${id}`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ category_id: categoryId }) })))
      setBatchCategoryId(''); setMessage(`已完成 ${selectedIds.length} 个文件的批量分类`); await refresh()
    } catch (e) { setError(e.message) } finally { setBatchUpdating(false) }
  }
  const openRename = () => {
    if (!selectedIds.length) return
    setRenameDraft({ find_text: '', replace_text: '', prefix: '', suffix: '', extension: '' })
    setRenameOpen(true)
  }
  const beginInlineRename = (doc) => {
    setInlineRenameId(doc.id)
    setInlineRenameValue(doc.original_name || doc.name)
    setError('')
  }
  const submitInlineRename = async (doc, targetName = inlineRenameValue, recordUndo = true) => {
    const nextName = String(targetName || '').trim()
    const currentName = doc.original_name || doc.name
    if (!nextName || nextName === currentName) { setInlineRenameId(null); return }
    const splitName = value => {
      const dot = value.lastIndexOf('.')
      return dot > 0 ? { stem: value.slice(0, dot), extension: value.slice(dot) } : { stem: value, extension: null }
    }
    const previous = splitName(currentName)
    const next = splitName(nextName)
    try {
      setInlineRenameBusy(true); setError('')
      await request('/documents/batch/rename', { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ids: [doc.id], find_text: previous.stem, replace_text: next.stem, extension: next.extension }) })
      setInlineRenameId(null)
      if (recordUndo) setLastRenameUndo({ id: doc.id, previousName: currentName, nextName })
      setMessage(`文件名已改为“${nextName}”${recordUndo ? '，按 Ctrl+Z 可撤回' : ''}`)
      await refresh()
    } catch (e) { setError(e.message) } finally { setInlineRenameBusy(false) }
  }
  const undoInlineRename = async () => {
    if (!lastRenameUndo || inlineRenameBusy) return
    const doc = documents.find(item => item.id === lastRenameUndo.id)
    if (!doc) return
    setLastRenameUndo(null)
    await submitInlineRename(doc, lastRenameUndo.previousName, false)
    setMessage(`已撤回文件名修改，恢复为“${lastRenameUndo.previousName}”。`)
  }
  useEffect(() => {
    const handleUndo = event => {
      if (!(event.ctrlKey || event.metaKey) || event.key.toLowerCase() !== 'z' || inlineRenameId || !lastRenameUndo) return
      const target = event.target
      if (target && ['INPUT', 'TEXTAREA', 'SELECT'].includes(target.tagName)) return
      event.preventDefault()
      undoInlineRename()
    }
    window.addEventListener('keydown', handleUndo)
    return () => window.removeEventListener('keydown', handleUndo)
  }, [inlineRenameId, lastRenameUndo, inlineRenameBusy, documents])
  const renderDocumentName = doc => inlineRenameId === doc.id
    ? <input className="inline-name-input" autoFocus value={inlineRenameValue} disabled={inlineRenameBusy} onChange={event => setInlineRenameValue(event.target.value)} onBlur={() => submitInlineRename(doc)} onKeyDown={event => { if (event.key === 'Enter') event.currentTarget.blur(); if (event.key === 'Escape') setInlineRenameId(null) }} aria-label={'修改 ' + doc.name}/>
    : <strong className="file-name-edit" role="button" tabIndex="0" title="单击直接修改文件名" onClick={() => beginInlineRename(doc)} onKeyDown={event => { if (event.key === 'Enter' || event.key === ' ') beginInlineRename(doc) }}>{doc.name}</strong>
  const renameSelected = async () => {
    if (!selectedIds.length || renaming) return
    try {
      setRenaming(true); setError('')
      const result = await request('/documents/batch/rename', { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ids: selectedIds, ...renameDraft, extension: renameDraft.extension.trim() || null }) })
      setMessage(`已批量修改 ${result.updated?.length || selectedIds.length} 个文件名；索引将自动更新。`)
      setSelectedIds([]); closeRename(); await refresh()
    } catch (e) { setError(e.message) } finally { setRenaming(false) }
  }
  const shutdownServices = () => {
    setStoppingServices(true)
    setError('')
    setMessage('已收到关闭服务操作。请双击项目根目录的“关闭平台.bat”完成关闭；数据会保留在 runtime 目录。')
    // Keep the button responsive while the user switches to the local BAT.
    window.setTimeout(() => setStoppingServices(false), 500)
  }
  const removeSelected = async () => {
    if (!selectedIds.length) return
    const confirmed = window.confirm(`确定永久移除选中的 ${selectedIds.length} 个文件吗？原文件、索引和数据库记录都会删除，无法恢复。`)
    if (!confirmed) return
    try {
      setError(''); setMessage(`正在移除 ${selectedIds.length} 个文件…`)
      const result = await request('/documents/batch', { method: 'DELETE', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ids: selectedIds }) })
      setMessage(`已移除 ${result.removed} 个文件${result.missing ? `，${result.missing} 个文件已不存在` : ''}`)
      await refresh()
    } catch (e) { setError(e.message) }
  }
  const formatOptions = useMemo(() => [...new Set(documents.map(doc => doc.extension.toLowerCase()))].sort(), [documents])
  const filteredDocuments = useMemo(() => documents.filter(doc => {
    const extension = doc.extension.toLowerCase()
    if (formatFilters.length && !formatFilters.includes(extension)) return false
    const timestamp = new Date(doc.uploaded_at).getTime()
    if (dateFrom && timestamp < new Date(`${dateFrom}T00:00:00`).getTime()) return false
    if (dateTo && timestamp > new Date(`${dateTo}T23:59:59.999`).getTime()) return false
    return true
  }), [documents, formatFilters, dateFrom, dateTo])
  const visibleDocuments = useMemo(() => [...filteredDocuments].sort((a, b) => {
    if (sortMethod === 'name_asc' || sortMethod === 'name_desc') {
      const result = a.name.localeCompare(b.name, 'zh-Hans', { numeric: true, sensitivity: 'base' })
      return sortMethod === 'name_asc' ? result : -result
    }
    if (sortMethod === 'size_asc' || sortMethod === 'size_desc') {
      const result = a.size_bytes - b.size_bytes
      return sortMethod === 'size_asc' ? result : -result
    }
    if (sortMethod === 'category_asc' || sortMethod === 'category_desc') {
      const result = categoryPath(a.category_id).localeCompare(categoryPath(b.category_id), 'zh-Hans', { sensitivity: 'base' })
      return sortMethod === 'category_asc' ? result : -result
    }
    if (sortMethod === 'status_asc' || sortMethod === 'status_desc') {
      const result = (a.index_status || '').localeCompare(b.index_status || '', 'en', { sensitivity: 'base' })
      return sortMethod === 'status_asc' ? result : -result
    }
    const result = new Date(a.uploaded_at).getTime() - new Date(b.uploaded_at).getTime()
    return sortMethod === 'modified_asc' ? result : -result
  }), [filteredDocuments, sortMethod, categories])
  const totalPages = pageSize > 0 ? Math.max(1, Math.ceil(visibleDocuments.length / pageSize)) : 1
  useEffect(() => { setCurrentPage(1) }, [archived, category, formatFilters, dateFrom, dateTo, sortMethod, pageSize])
  useEffect(() => { if (currentPage > totalPages) setCurrentPage(totalPages) }, [currentPage, totalPages])
  const pagedDocuments = useMemo(() => {
    if (pageSize <= 0) return visibleDocuments
    const start = (currentPage - 1) * pageSize
    return visibleDocuments.slice(start, start + pageSize)
  }, [visibleDocuments, pageSize, currentPage])
  // Build the tree from the already filtered/sorted rows so format, date,
  // archive, category and sort controls stay consistent in either view.
  const categoryTree = useMemo(() => {
    const groups = new Map(categories.map(item => [String(item.id), { ...item, documents: [], children: [] }]))
    const roots = []
    groups.forEach(group => {
      const parent = group.parent_id == null ? null : groups.get(String(group.parent_id))
      if (parent) parent.children.push(group); else roots.push(group)
    })
    const uncategorized = { id: null, name: '未分类', documents: [], children: [] }
    roots.unshift(uncategorized)
    visibleDocuments.forEach(doc => {
      const group = doc.category_id == null ? uncategorized : groups.get(String(doc.category_id))
      if (group) group.documents.push(doc)
    })
    const sortTree = nodes => nodes.sort((a, b) => a.name.localeCompare(b.name, 'zh-Hans')).forEach(node => sortTree(node.children))
    sortTree(roots)
    return roots
  }, [categories, pagedDocuments])
  const toggleCategoryNode = (id) => setExpandedCategoryIds(previous => {
    const next = new Set(previous)
    const key = String(id)
    if (next.has(key)) next.delete(key); else next.add(key)
    return next
  })
  const openCategoryRename = (item) => { beginRenameCategory(item); setCategoryOpen(true) }
  const countTreeDocuments = (group) => group.documents.length + group.children.reduce((total, child) => total + countTreeDocuments(child), 0)
  const renderCategoryNode = (group) => {
    const key = group.id == null ? '__uncategorized' : String(group.id)
    const expanded = expandedCategoryIds.has(key)
    return <div className="tree-group" key={key}><div className="tree-node"><button className="tree-expander" onClick={() => toggleCategoryNode(key)} aria-label={expanded ? '收起分类' : '展开分类'}>{expanded ? '▾' : '▸'}</button><strong>{group.name}</strong><span className="tree-count">{countTreeDocuments(group)} 个文件</span>{group.id != null && <><button className="text-button" onClick={() => openCategoryRename(group)}>改名</button><button className="text-button" onClick={() => setCategory(String(group.id))}>筛选</button></>}</div>{expanded && group.documents.length > 0 && <div className="tree-children">{group.documents.map(doc => <div className="tree-file" key={doc.id}><input type="checkbox" checked={selectedIds.includes(doc.id)} onChange={() => toggleSelected(doc.id)} aria-label={'选择 ' + doc.name}/><div className="tree-file-main">{renderDocumentName(doc)}<small>{doc.extension.toUpperCase()} · {formatBytes(doc.size_bytes)} · {new Date(doc.uploaded_at).toLocaleString()}</small></div><select value={doc.category_id || ''} onChange={e => update(doc, { category_id: e.target.value ? Number(e.target.value) : null }, true)} aria-label={'修改 ' + doc.name + ' 的分类'}><option value="">未分类</option>{orderedCategories.map(({ item: c, depth }) => <option key={c.id} value={c.id}>{categoryPath(c.id)}</option>)}</select><Status value={doc.index_status}/><div className="tree-file-actions"><button onClick={() => openPreview(doc)}>预览</button><a href={API + '/documents/' + doc.id + '/download'}>下载</a>{doc.index_status === 'failed' && <button onClick={() => retry(doc)}>重试</button>}<button onClick={() => update(doc, { archived: !archived }, true)}>{archived ? '恢复' : '归档'}</button></div></div>)}</div>}{expanded && group.children.length > 0 && <div className="tree-children tree-nested">{group.children.map(renderCategoryNode)}</div>}{expanded && group.documents.length === 0 && group.children.length === 0 && <div className="tree-empty">此分类暂无文件</div>}</div>
  }
  const toggleSelected = (id) => setSelectedIds(ids => ids.includes(id) ? ids.filter(item => item !== id) : [...ids, id])
  const toggleAll = () => {
    const visibleIds = pagedDocuments.map(doc => doc.id)
    const allVisibleSelected = visibleIds.length > 0 && visibleIds.every(id => selectedIds.includes(id))
    setSelectedIds(ids => allVisibleSelected ? ids.filter(id => !visibleIds.includes(id)) : [...new Set([...ids, ...visibleIds])])
  }
  const updateSortMethod = async (value) => {
    setSortMethod(value)
    try { await request('/settings', { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ sort_method: value }) }) } catch (e) { setError(e.message) }
  }
  const toggleColumnSort = (column) => {
    const current = sortMethod === `${column}_asc` ? 'asc' : sortMethod === `${column}_desc` ? 'desc' : null
    updateSortMethod(`${column}_${current === 'asc' ? 'desc' : 'asc'}`)
  }
  const sortIndicator = column => sortMethod === `${column}_asc` ? ' ↑' : sortMethod === `${column}_desc` ? ' ↓' : ''
  const toggleFormatFilter = (extension) => setFormatFilters(filters => filters.includes(extension) ? filters.filter(item => item !== extension) : [...filters, extension])
  const selectFilteredByFormat = () => setSelectedIds(ids => [...new Set([...ids, ...visibleDocuments.map(doc => doc.id)])])
  const clearFileFilters = () => {
    setFormatFilters([])
    setDateFrom('')
    setDateTo('')
    // 清除筛选时同步取消“选中这些格式”带来的文件选择，避免筛选条件已清空但列表仍显示勾选。
    setSelectedIds([])
  }
  const downloadSelected = () => {
    if (!selectedIds.length) return
    const params = new URLSearchParams(); selectedIds.forEach(id => params.append('ids', id))
    window.location.href = `${API}/documents/download-batch?${params}`
  }
  const visibleResults = useMemo(() => {
    const all = results?.results || []
    if (!results || results.mode !== 'semantic' || showAllResults) return all
    return all.slice(0, Math.max(1, semanticResultLimit))
  }, [results, showAllResults, semanticResultLimit])
  const hiddenResultCount = results?.mode === 'semantic' ? Math.max(0, (results.results?.length || 0) - Math.max(1, semanticResultLimit)) : 0
  const fileFilterPanel = <div className="library-filter-panel"><div className="file-filter-head"><div><strong>筛选资料</strong><small>按文件格式和修改时间范围过滤当前资料列表。</small></div><button className="ghost" onClick={clearFileFilters} disabled={!formatFilters.length && !dateFrom && !dateTo && !selectedIds.length}>清除筛选</button></div><div className="file-filter-controls"><div className="format-filter"><span className="filter-label">文件格式</span><div className="format-checks">{formatOptions.length ? formatOptions.map(extension => <label className="format-check" key={extension}><input type="checkbox" checked={formatFilters.includes(extension)} onChange={() => toggleFormatFilter(extension)}/><span>{extension.replace('.', '').toUpperCase()}</span></label>) : <small>暂无可筛选格式</small>}</div><button className="text-button select-format-button" onClick={selectFilteredByFormat} disabled={!formatFilters.length || !visibleDocuments.length}>选中这些格式</button></div><div className="date-filter"><span className="filter-label">修改时间</span><label>从 <input type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)} /></label><label>到 <input type="date" value={dateTo} onChange={e => setDateTo(e.target.value)} /></label></div></div></div>
  const paginationControls = pageSize > 0 && visibleDocuments.length > 0 && <div className="pagination-controls"><span>第 {currentPage} / {totalPages} 页 · 共 {visibleDocuments.length} 个文件</span><div><button className="ghost" onClick={() => setCurrentPage(page => Math.max(1, page - 1))} disabled={currentPage <= 1}>上一页</button><button className="ghost" onClick={() => setCurrentPage(page => Math.min(totalPages, page + 1))} disabled={currentPage >= totalPages}>下一页</button></div></div>
  const renderCategoryRow = ({ item, depth }) => {
    const parentOptions = orderedCategories.filter(({ item: option }) => option.id !== item.id)
    return <tr key={item.id}><td>{editingCategoryId === item.id ? <><input className="category-edit-input" value={editingCategoryName} onChange={e => setEditingCategoryName(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') saveRenameCategory(item); if (e.key === 'Escape') cancelRenameCategory() }} aria-label={'修改分类 ' + item.name}/><select className="category-parent-select category-edit-parent" value={editingCategoryParentId} onChange={e => setEditingCategoryParentId(e.target.value)} aria-label={'修改父分类 ' + item.name}><option value="">顶级分类</option>{parentOptions.map(({ item: option, depth: optionDepth }) => <option key={option.id} value={option.id}>{categoryPath(option.id)}</option>)}</select></> : <strong>{categoryPath(item.id)}</strong>}</td><td>{item.parent_id ? categoryPath(item.parent_id) : '顶级分类'}</td><td>{documents.filter(doc => doc.category_id === item.id).length}</td><td className="category-actions">{editingCategoryId === item.id ? <><button className="text-button" onClick={() => saveRenameCategory(item)} disabled={categoryBusy}>保存</button><button className="text-button muted-button" onClick={cancelRenameCategory} disabled={categoryBusy}>取消</button></> : <><button className="text-button" onClick={() => beginRenameCategory(item)} disabled={categoryBusy}>修改名称</button><button className="text-button danger-text" onClick={() => deleteCategory(item)} disabled={categoryBusy}>删除</button></>}</td></tr>
  }
  const remoteEmbeddingFields = settingsDraft && settingsDraft.embedding_provider !== 'local' && <><label><strong>模型接口地址</strong><input value={settingsDraft.embedding_api_url} onChange={e => setSettingsDraft({ ...settingsDraft, embedding_api_url: e.target.value })} placeholder="https://.../embeddings" /></label><label><strong>模型名称</strong><input value={settingsDraft.embedding_model} onChange={e => setSettingsDraft({ ...settingsDraft, embedding_model: e.target.value })} placeholder="可选" /></label><label><strong>API Key</strong><input type="password" value={settingsDraft.embedding_api_key} onChange={e => setSettingsDraft({ ...settingsDraft, embedding_api_key: e.target.value })} placeholder={settingsDraft.embedding_api_key_set ? '已设置，留空表示保持不变' : '未设置'} /></label></>

  const workspaceOpacity = Number.isFinite(Number(themeSettings.workspace_opacity)) ? Math.min(1, Math.max(0.55, Number(themeSettings.workspace_opacity))) : 0.96
  const backgroundOpacity = Number.isFinite(Number(themeSettings.background_opacity)) ? Math.min(1, Math.max(0, Number(themeSettings.background_opacity))) : 0.38
  const accentOpacity = Number.isFinite(Number(themeSettings.accent_opacity)) ? Math.min(1, Math.max(0, Number(themeSettings.accent_opacity))) : 1
  const shellStyle = { '--workspace-opacity': String(workspaceOpacity), '--accent-color': hexToRgba(themeSettings.accent_color, accentOpacity), ...(backgroundUrl ? { backgroundImage: `linear-gradient(rgba(255,255,255,${1 - backgroundOpacity}), rgba(255,255,255,${1 - backgroundOpacity})), url(${JSON.stringify(backgroundUrl)})`, backgroundSize: 'cover', backgroundPosition: 'center', backgroundAttachment: 'fixed' } : {}) }
  return <div className={`app-shell${backgroundUrl ? ' has-custom-background' : ''}`} data-modal-animation={themeSettings.modal_animation || 'fade'} style={shellStyle}>
     <header className="topbar"><div><div className="eyebrow" style={{ color: themeSettings.eyebrow_color }}>{themeSettings.eyebrow_text}</div><h1 style={{ color: themeSettings.title_color }}>{themeSettings.title_text}</h1><p className="creator-line" style={{ color: themeSettings.subtitle_color }}>{themeSettings.subtitle_text}</p></div><div className="topbar-actions"><button className="ghost topbar-ghost" onClick={openSettings}>设置</button><button className="ghost topbar-ghost" onClick={openDatabase}>查看数据库</button></div></header>
    <main>
       <section className="hero panel"><div><h2>资料工作台</h2><p>支持 PDF、DOCX、PPT/PPTX、XLS/XLSX、TXT、Markdown、CSV、JSON、YAML、XML、HTML、代码和配置文件；PPTX、XLSX、XLS 可直接预览原文；可一次选择多个文件。</p><small className="storage-notice">浏览器无删除全局文件权限，平台只会复制一份到数据库中，源文件保持不变。</small>{storageInfo && <small className="storage-policy">数据库目录：{storageInfo.database_host_path || './runtime/postgres'} · 文件目录：{storageInfo.host_path || storageInfo.storage_dir}</small>}</div><div className="upload-actions"><label className="upload primary">{uploading ? '批量上传中…' : '批量上传'}<input type="file" multiple accept=".pdf,.docx,.ppt,.pptx,.xls,.xlsx,.txt,.md,.markdown,.log,.csv,.tsv,.json,.yaml,.yml,.xml,.html,.htm,.css,.js,.jsx,.ts,.tsx,.py,.java,.go,.rs,.sql,.ini,.conf,.toml,.vue,.svelte" onChange={upload} disabled={uploading}/></label></div></section>

      {(message || error) && !categoryOpen && <div className={error ? 'notice error' : 'notice'}><button className="notice-dismiss" type="button" onClick={() => { setError(''); setMessage('') }} aria-label="关闭提示">×</button><span>{error || message}</span></div>}
      <section className="toolbar panel"><div className="search"><form onSubmit={search}><input value={query} onChange={e => setQuery(e.target.value)} placeholder={mode === 'semantic' ? '例如：同一请求重发会不会生成两份记录？' : mode === 'filename' ? '搜索文件名' : '搜索正文关键词'} /><select value={mode} onChange={e => setMode(e.target.value)} aria-label="搜索范围"><option value="filename">文件名</option><option value="content">关键词</option><option value="semantic">语义检索</option></select><button className="primary">搜索</button></form></div><div className="filters"><select value={category} onChange={e => setCategory(e.target.value)}><option value="">全部分类</option>{orderedCategories.map(({ item: c, depth }) => <option key={c.id} value={c.id}>{'　'.repeat(depth)}{c.name}</option>)}</select></div></section>
       {results && <section className="panel results"><div className="section-heading"><div><h2>检索结果</h2><span>{results.results.length} 个结果 · {results.mode === 'semantic' ? '语义检索' : results.mode === 'filename' ? '文件名' : '关键词'}</span></div><button className="ghost" onClick={() => setResults(null)}>×</button></div>{visibleResults.length === 0 ? <div className="empty">没有找到匹配资料{!archived && <><br/><small>默认检索会排除归档文件；可切换到“归档区”后再搜索。</small></>}</div> : visibleResults.map(r => { const snippet = resultSnippet(r); return <article className="result" key={r.document.id}><div><h3><Highlight text={r.document.name} query={results.query}/></h3><div className="meta">{categoryPath(r.document.category_id)} · {r.document.extension.toUpperCase()} {r.score != null && `· 相关度 ${r.score}`}</div></div>{snippet && <p><Highlight text={snippet} query={results.query}/></p>}<div className="result-actions"><button className="text-button" onClick={() => openPreview(r.document)}>预览</button><a href={`${API}/documents/${r.document.id}/download`}>下载原文件</a></div></article> })}{results.mode === 'semantic' && hiddenResultCount > 0 && <div className="result-more"><button className="ghost" onClick={() => setShowAllResults(value => !value)}>{showAllResults ? '收起其他结果' : `显示其他 ${hiddenResultCount} 个文件`}</button></div>}</section>}
       <section className={`panel library ${libraryView === 'tree' ? 'tree-mode' : ''}`}><div className="section-heading"><div><h2>{archived ? '归档资料' : '资料列表'}</h2><span>{visibleDocuments.length} 个可见文件{visibleDocuments.length !== documents.length ? '（共 ' + documents.length + ' 个）' : ''}{selectedIds.length ? ' · 已选 ' + selectedIds.length + ' 个' : ''}{hasIndexing ? ' · 正在自动更新索引状态' : ''}</span></div><div className="library-actions"><div className="library-view-toggle" role="group" aria-label="资料显示方式"><button className={libraryView === 'list' ? 'toggle active' : 'toggle'} onClick={() => setLibraryView('list')}>列表</button><button className={libraryView === 'tree' ? 'toggle active' : 'toggle'} onClick={() => setLibraryView('tree')}>分类树</button></div><button className="ghost category-library-button" onClick={() => setCategoryOpen(true)}>分类管理</button><button className={archived ? 'toggle active' : 'toggle'} onClick={() => setArchived(v => !v)}>{archived ? '归档区' : '默认资料'}</button>{selectedIds.length > 0 && <div className="batch-category-control"><select value={batchCategoryId} onChange={e => setBatchCategoryId(e.target.value)} aria-label="批量分类"><option value="">移到未分类</option>{orderedCategories.map(({ item: c, depth }) => <option key={c.id} value={c.id}>移到：{categoryPath(c.id)}</option>)} </select><button className="ghost" onClick={batchCategorize} disabled={batchUpdating}>{batchUpdating ? '分类中…' : '批量分类'}</button></div>}<button className="ghost" onClick={openRename} disabled={!selectedIds.length} title={selectedIds.length ? `批量修改已选 ${selectedIds.length} 个文件` : '请先勾选文件'}>批量改名</button><button className="ghost" onClick={downloadSelected} disabled={!selectedIds.length}>批量下载 ZIP</button><button className="danger-button" onClick={removeSelected} disabled={!selectedIds.length}>批量移除</button><button className="ghost" onClick={scanStorage} disabled={scanningStorage}>{scanningStorage ? '扫描中…' : '扫描文件状态'}</button><button className="ghost" onClick={() => refresh()}>刷新</button></div></div>{fileFilterPanel}{documents.length === 0 ? <div className="empty">这里还没有资料。请上传支持的文件类型。</div> : visibleDocuments.length === 0 ? <div className="empty">当前筛选条件下没有文件。可以清除筛选后重试。</div> : <div className="table-wrap"><table><thead><tr><th className="check-cell"><input type="checkbox" checked={pagedDocuments.length > 0 && pagedDocuments.every(doc => selectedIds.includes(doc.id))} onChange={toggleAll} aria-label="全选当前筛选结果"/></th><th><button className="table-sort-button" onClick={() => toggleColumnSort('name')}>名称{sortIndicator('name')}</button></th><th><button className="table-sort-button" onClick={() => toggleColumnSort('category')}>分类{sortIndicator('category')}</button></th><th><button className="table-sort-button" onClick={() => toggleColumnSort('size')}>大小{sortIndicator('size')}</button></th><th><button className="table-sort-button" onClick={() => toggleColumnSort('status')}>索引状态{sortIndicator('status')}</button></th><th><button className="table-sort-button" onClick={() => toggleColumnSort('modified')}>上传时间{sortIndicator('modified')}</button></th><th></th></tr></thead><tbody>{pagedDocuments.map(doc => <tr key={doc.id}><td className="check-cell"><input type="checkbox" checked={selectedIds.includes(doc.id)} onChange={() => toggleSelected(doc.id)} aria-label={'选择 ' + doc.name}/></td><td>{renderDocumentName(doc)}<small>{doc.extension.toUpperCase()}</small></td><td><select value={doc.category_id || ''} onChange={e => update(doc, { category_id: e.target.value ? Number(e.target.value) : null })}><option value="">未分类</option>{orderedCategories.map(({ item: c, depth }) => <option key={c.id} value={c.id}>{categoryPath(c.id)}</option>)} </select></td><td>{formatBytes(doc.size_bytes)}</td><td><StorageStatus value={doc.storage_status}/><Status value={doc.index_status}/>{doc.index_error && <small className="error-text">{doc.index_error}</small>}</td><td>{new Date(doc.uploaded_at).toLocaleString()}</td><td className="actions"><button onClick={() => openPreview(doc)}>预览</button><a href={API + '/documents/' + doc.id + '/download'}>下载</a>{doc.index_status === 'failed' && <button onClick={() => retry(doc)}>重试</button>}{doc.storage_status === 'missing' && <button onClick={() => deleteMissingDocument(doc)}>删除文件</button>}<button onClick={() => update(doc, { archived: !archived })}>{archived ? '恢复' : '归档'}</button></td></tr>)}</tbody></table></div>}{paginationControls}</section>
{libraryView === 'tree' && <section className="panel category-tree-panel"><div className="tree-help">按分类层级展开文件；文件分类、归档和索引操作可直接修改。</div>{categoryTree.length === 0 ? <div className="empty">当前没有分类或文件。</div> : <div className="category-tree">{categoryTree.map(renderCategoryNode)}</div>}</section>}
    </main><footer><span className="creator-mark">KOKONA 编写制作</span></footer>
    {categoryOpen && <div className={`modal-backdrop ${closingModal ? 'modal-closing' : ''}`} role="presentation" onClick={event => { if (event.target === event.currentTarget) closeCategory() }}><section className="preview-modal category-modal" role="dialog" aria-modal="true" aria-label="分类管理"><div className="section-heading"><div><h2>分类管理</h2><span>直接添加、修改或删除分类；删除分类时文件会回到未分类。</span></div><button className="ghost" onClick={() => closeCategory()}>×</button></div>{(error || message) && <div className={error ? 'modal-inline-error' : 'modal-inline-message'}><button className="notice-dismiss" type="button" onClick={() => { setError(''); setMessage('') }} aria-label="关闭提示">×</button><span>{error || message}</span></div>}<div className="category-create-row"><input value={newCategoryName} onChange={e => setNewCategoryName(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') createCategory() }} placeholder="输入新分类名称" aria-label="新分类名称"/><select className="category-parent-select" value={newCategoryParentId} onChange={e => setNewCategoryParentId(e.target.value)} aria-label="父分类"><option value="">顶级分类</option>{orderedCategories.map(({ item, depth }) => <option key={item.id} value={item.id}>{categoryPath(item.id)}</option>)}</select><button className="primary" onClick={createCategory} disabled={categoryBusy || !newCategoryName.trim()}>添加分类</button></div>{categories.length === 0 ? <div className="empty">还没有分类，先在上方输入名称并添加。</div> : <div className="table-wrap"><table className="category-table"><thead><tr><th>分类层级</th><th>父分类</th><th>当前列表文件</th><th>操作</th></tr></thead><tbody>{orderedCategories.map(renderCategoryRow)}</tbody></table></div>}</section></div>}
    {preview && <div className={`modal-backdrop ${closingModal ? 'modal-closing' : ''}`} role="presentation" onClick={event => { if (event.target === event.currentTarget) closePreview() }}><section className="preview-modal" role="dialog" aria-modal="true" aria-label="文件预览"><div className="section-heading"><div><h2>{preview.document.name}</h2><span>{preview.document.extension.toUpperCase()} · <Status value={preview.document.index_status}/></span></div><button className="ghost" onClick={() => closePreview()}>×</button></div>{preview.rich_available && <div className="preview-switch" role="group" aria-label="预览模式"><button className={previewMode === 'rich' ? 'active' : ''} onClick={() => setPreviewMode('rich')}>原文视图</button><button className={previewMode === 'plain' ? 'active' : ''} onClick={() => setPreviewMode('plain')}>纯文本</button></div>}{previewMode === 'rich' && preview.rich_available ? (preview.rich_type === 'pdf' ? <PdfPreview src={API + '/documents/' + preview.document.id + '/preview-data'} /> : preview.rich_type === 'iframe' ? <iframe className="rich-preview-frame" src={API + '/documents/' + preview.document.id + '/preview-file'} title="原文预览" sandbox="" /> : preview.rich_type === 'pptx' ? <PptxPreview src={API + '/documents/' + preview.document.id + '/preview-file'} fallbackHtml={preview.rich_html} /> : preview.rich_type === 'sheet' ? <SpreadsheetPreview src={API + '/documents/' + preview.document.id + '/preview-file'} fallbackHtml={preview.rich_html} /> : <div className="rich-preview" dangerouslySetInnerHTML={{ __html: preview.rich_html || '' }} />) : preview.previewable ? <pre className="preview-content"><Highlight text={preview.text} query={query}/></pre> : <div className="empty">{preview.reason || '暂时无法预览该文件'}</div>}{preview.truncated && <small className="preview-note">预览已截断，原文件可下载查看完整内容。</small>}</section></div>}
    {databaseOverview && <div className={`modal-backdrop ${closingModal ? 'modal-closing' : ''}`} role="presentation" onClick={event => { if (event.target === event.currentTarget) closeDatabase() }}><section className="preview-modal database-modal" role="dialog" aria-modal="true" aria-label="数据库概览"><div className="section-heading"><div><h2>数据库概览</h2></div><button className="ghost" onClick={() => closeDatabase()}>×</button></div><div className="database-summary"><div className="database-stat"><span>文件总数</span><strong>{databaseOverview.documents.length}</strong></div><div className="database-stat"><span>分类数量</span><strong>{databaseOverview.categories.length}</strong></div><div className="database-stat"><span>可检索</span><strong>{databaseOverview.documents.filter(doc => doc.index_status === 'ready').length}</strong></div><div className="database-stat"><span>归档文件</span><strong>{databaseOverview.documents.filter(doc => doc.archived).length}</strong></div></div><div className="table-wrap database-table-wrap"><table className="database-table"><colgroup><col className="database-col-file"/><col className="database-col-path"/><col className="database-col-category"/><col className="database-col-index"/><col className="database-col-archive"/></colgroup><thead><tr><th>文件</th><th>原文件路径</th><th>分类</th><th>索引状态</th><th>归档</th></tr></thead><tbody>{databaseOverview.documents.map(doc => <tr key={doc.id}><td className="database-file-cell"><strong>{doc.name}</strong><small>{doc.extension.toUpperCase()} · {formatBytes(doc.size_bytes)}</small></td><td className="database-path-cell"><code>{doc.storage_path || '未记录'}</code></td><td className="database-category-cell">{categoryPath(doc.category_id)}</td><td className="database-index-cell"><Status value={doc.index_status}/></td><td className="database-archive-cell">{doc.archived ? '是' : '否'}</td></tr>)}</tbody></table></div>{databaseOverview.documents.length === 0 && <div className="empty">数据库中还没有文件记录。</div>}</section></div>}
    {renameOpen && <div className={`modal-backdrop ${closingModal ? 'modal-closing' : ''}`} role="presentation" onClick={event => { if (event.target === event.currentTarget) closeRename() }}><section className="preview-modal rename-modal" role="dialog" aria-modal="true" aria-label="批量修改文件名"><div className="section-heading"><div><h2>批量修改文件名</h2><span>已选择 {selectedIds.length} 个文件；修改后会自动重新建立索引。</span></div><button className="ghost" aria-label="关闭" onClick={closeRename}>×</button></div><div className="rename-form"><label><strong>查找文件名片段</strong><input value={renameDraft.find_text} onChange={e => setRenameDraft({ ...renameDraft, find_text: e.target.value })} placeholder="例如：2024" /><small>只替换文件名主体，不修改原文件内容。</small></label><label><strong>替换为</strong><input value={renameDraft.replace_text} onChange={e => setRenameDraft({ ...renameDraft, replace_text: e.target.value })} placeholder="留空表示删除查找内容" /></label><label><strong>统一添加前缀</strong><input value={renameDraft.prefix} onChange={e => setRenameDraft({ ...renameDraft, prefix: e.target.value })} placeholder="可选，例如：已整理_" /></label><label><strong>统一添加后缀</strong><input value={renameDraft.suffix} onChange={e => setRenameDraft({ ...renameDraft, suffix: e.target.value })} placeholder="可选，例如：_复核" /></label><label><strong>统一修改扩展名</strong><input value={renameDraft.extension} onChange={e => setRenameDraft({ ...renameDraft, extension: e.target.value })} placeholder="可选，例如：.pdf 或 pdf" /><small>只修改文件名后缀，不会转换文件格式；请输入平台支持的扩展名。</small></label></div><div className="rename-warning">扩展名修改不会改变文件内容。若内容格式不匹配，预览和索引可能失败，但原文件仍可下载。</div><div className="settings-actions"><button className="ghost" onClick={closeRename}>取消</button><button className="primary" onClick={renameSelected} disabled={renaming}>{renaming ? '修改中…' : '应用修改'}</button></div></section></div>}
    {settingsOpen && settingsDraft && <div className={`modal-backdrop ${closingModal ? 'modal-closing' : ''}`} role="presentation" onClick={event => { if (event.target === event.currentTarget) closeSettings() }}><section className="preview-modal settings-modal" role="dialog" aria-modal="true" aria-label="系统设置"><div className="section-heading"><div><h2>系统设置</h2><span>设置保存到数据库，立即作用于新上传和索引任务</span></div><button className="ghost" onClick={() => closeSettings()}>×</button></div><div className="settings-grid"><div className="settings-group-title">文件操作</div><label className="settings-path"><strong>数据库存放目录</strong><input value={settingsDraft.database_host_path || ''} onChange={e => setSettingsDraft({ ...settingsDraft, database_host_path: e.target.value })} placeholder="./runtime/postgres 或 E:/data/postgres" /><small>确认保存后会停止服务，移动 PostgreSQL 数据目录并自动重启；目标目录必须为空。</small></label><label className="settings-path"><strong>文档存放目录</strong><input value={settingsDraft.storage_host_path || ''} onChange={e => setSettingsDraft({ ...settingsDraft, storage_host_path: e.target.value })} placeholder="./runtime/storage 或 E:/data/storage" /><small>确认保存后会移动已上传文档和分类目录；原目录迁移成功后才会清理。</small></label><label><strong>文件存储方式</strong><select value={settingsDraft.storage_layout} onChange={e => setSettingsDraft({ ...settingsDraft, storage_layout: e.target.value })}><option value="flat">统一目录（源文件名，分类只改数据库）</option><option value="folders">按源文件名分目录（分类/归档同步移动平台文件）</option></select><small>分类和归档只会在平台存储目录内移动文件，不会删除本地原文件。</small></label><label><strong>语义模型</strong><select value={settingsDraft.embedding_provider} onChange={e => setSettingsDraft({ ...settingsDraft, embedding_provider: e.target.value })}><option value="local">本地向量（无需联网）</option><option value="remote">联网模型（OpenAI-compatible）</option></select></label>{remoteEmbeddingFields}<label><strong>资料列表分页</strong><select value={settingsDraft.page_size ?? 20} onChange={e => setSettingsDraft({ ...settingsDraft, page_size: Number(e.target.value) })}><option value="10">每页 10 个</option><option value="20">每页 20 个</option><option value="50">每页 50 个</option><option value="100">每页 100 个</option><option value="200">每页 200 个</option><option value="0">不分页</option></select><small>默认分页；选择“不分页”会一次显示当前筛选后的全部文件。</small></label><label><strong>语义检索首屏显示数量</strong><input type="number" min="1" max="50" value={settingsDraft.semantic_result_limit ?? 5} onChange={e => setSettingsDraft({ ...settingsDraft, semantic_result_limit: Number(e.target.value) })}/><small>自然语言检索默认展开的文件数，其余结果会折叠。</small></label><label><strong>单文件上限（MB）</strong><input type="number" min="1" value={settingsDraft.max_file_mb} onChange={e => setSettingsDraft({ ...settingsDraft, max_file_mb: e.target.value })} /></label><label><strong>批量下载上限（GB）</strong><input type="number" min="0.1" step="0.1" value={settingsDraft.max_batch_download_gb} onChange={e => setSettingsDraft({ ...settingsDraft, max_batch_download_gb: e.target.value })} /></label><label><strong>索引轮询间隔（秒）</strong><input type="number" min="0.2" max="60" step="0.1" value={settingsDraft.index_poll_seconds} onChange={e => setSettingsDraft({ ...settingsDraft, index_poll_seconds: e.target.value })} /><small>后台 worker 检查新索引任务的间隔。</small></label><div className="settings-group-title">外观</div><label className="settings-path settings-background"><strong>操作台背景图片</strong><input type="file" accept=".png,.jpg,.jpeg,.webp,.gif,image/png,image/jpeg,image/webp,image/gif" onChange={uploadBackground} disabled={backgroundUploading} /><small>{backgroundUploading ? '正在上传背景图片…' : '支持 PNG、JPG、WEBP、GIF，单张不超过 10 MB；上传后立即生效并保存在 runtime/ui-background。'}</small>{backgroundUrl && <div className="settings-background-preview"><img src={backgroundUrl} alt="当前操作台背景" /><button className="ghost" type="button" onClick={clearBackground}>清除背景</button></div>}</label><label><strong>操作台透明度：{Math.round(Number(settingsDraft.workspace_opacity ?? 0.96) * 100)}%</strong><input type="range" min="0.55" max="1" step="0.01" value={settingsDraft.workspace_opacity ?? 0.96} onChange={e => setSettingsDraft({ ...settingsDraft, workspace_opacity: Number(e.target.value) })} /><small>调整资料列表、工作台面板的透明度。</small></label><label><strong>背景图片透明度：{Math.round(Number(settingsDraft.background_opacity ?? 0.38) * 100)}%</strong><input type="range" min="0" max="1" step="0.01" value={settingsDraft.background_opacity ?? 0.38} onChange={e => setSettingsDraft({ ...settingsDraft, background_opacity: Number(e.target.value) })} /><small>透明度越高，背景图片越清晰。</small></label><label className="settings-accent-opacity"><strong>衬色透明度：{Math.round(Number(settingsDraft.accent_opacity ?? 1) * 100)}%</strong><input type="range" min="0.35" max="1" step="0.01" value={settingsDraft.accent_opacity ?? 1} onChange={e => setSettingsDraft({ ...settingsDraft, accent_opacity: Number(e.target.value) })} /></label><label className="settings-color-control settings-accent-color"><strong>衬色颜色</strong><input type="color" value={settingsDraft.accent_color || "#173b45"} onChange={e => setSettingsDraft({ ...settingsDraft, accent_color: e.target.value })} /><small>用于顶部标题栏和主要强调区域。</small></label><label className="settings-eyebrow-text"><strong>顶部小标题文字</strong><input value={settingsDraft.eyebrow_text ?? "Documents Workspace"} onChange={e => setSettingsDraft({ ...settingsDraft, eyebrow_text: e.target.value })} maxLength={120} /><small>显示在主标题上方的 Documents Workspace 位置。</small></label><label className="settings-color-control settings-eyebrow-color"><strong>顶部小标题颜色</strong><input type="color" value={settingsDraft.eyebrow_color || "#9ad4c8"} onChange={e => setSettingsDraft({ ...settingsDraft, eyebrow_color: e.target.value })} /></label><label className="settings-title-text"><strong>主标题文字</strong><input value={settingsDraft.title_text ?? "文件管理与知识检索平台"} onChange={e => setSettingsDraft({ ...settingsDraft, title_text: e.target.value })} maxLength={120} /><small>显示在页面顶部的大标题位置。</small></label><label className="settings-color-control settings-title-color"><strong>主标题颜色</strong><input type="color" value={settingsDraft.title_color || "#f5fbfa"} onChange={e => setSettingsDraft({ ...settingsDraft, title_color: e.target.value })} /></label><label className="settings-subtitle-text"><strong>副标题文字</strong><input value={settingsDraft.subtitle_text ?? "Made By KOKONA"} onChange={e => setSettingsDraft({ ...settingsDraft, subtitle_text: e.target.value })} maxLength={120} /><small>显示在主标题下方，可留空。</small></label><label className="settings-color-control settings-subtitle-color"><strong>副标题颜色</strong><input type="color" value={settingsDraft.subtitle_color || "#9fc4c9"} onChange={e => setSettingsDraft({ ...settingsDraft, subtitle_color: e.target.value })} /></label><label><strong>悬浮窗动效</strong><select value={settingsDraft.modal_animation || "fade"} onChange={e => setSettingsDraft({ ...settingsDraft, modal_animation: e.target.value })}><option value="fade">淡入淡出</option><option value="scale">缩放淡入</option><option value="slide">向上滑入</option><option value="none">无动效</option></select><small>应用于设置、预览、数据库和分类管理窗口。</small></label></div>{['queued', 'processing'].includes(migrationStatus) && <div className="settings-migration-notice">存储迁移进行中：服务会暂时不可用，完成后页面会自动更新目录。</div>}{migrationStatus === 'failed' && <div className="settings-migration-error">上次存储迁移失败：{storageInfo?.migration?.message || '请查看 runtime/.storage-migration.log'}</div>}<div className="settings-actions"><button className="danger-button settings-shutdown" onClick={shutdownServices} disabled={stoppingServices}>{stoppingServices ? '正在处理…' : '结束服务'}</button><button className="ghost" onClick={() => closeSettings()}>取消</button><button className="primary" onClick={saveSettings} disabled={savingSettings || ['queued', 'processing'].includes(migrationStatus)}>{savingSettings ? '保存中…' : '保存设置'}</button></div></section></div>}
  </div>
}
