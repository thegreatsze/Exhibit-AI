import { execFile, exec } from 'child_process'
import { promisify } from 'util'
import path from 'path'
import fs from 'fs'
import os from 'os'
import { PDFDocument } from 'pdf-lib'
import sharp from 'sharp'
import { simpleParser } from 'mailparser'
import { BrowserWindow } from 'electron'

const execFileAsync = promisify(execFile)
const execAsync = promisify(exec)

export async function findLibreOffice(): Promise<string | null> {
  const candidates =
    process.platform === 'win32'
      ? [
          'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
          'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe'
        ]
      : process.platform === 'darwin'
        ? ['/Applications/LibreOffice.app/Contents/MacOS/soffice']
        : ['/usr/bin/soffice', '/usr/local/bin/soffice']

  for (const c of candidates) {
    if (fs.existsSync(c)) return c
  }
  // Try PATH
  try {
    const cmd = process.platform === 'win32' ? 'where soffice' : 'which soffice'
    const { stdout } = await execAsync(cmd)
    const found = stdout.trim().split('\n')[0]
    if (found) return found
  } catch {
    // not found on PATH
  }
  return null
}

export async function convertToPdf(
  inputPath: string,
  outputDir: string,
  outputName: string
): Promise<{ pdfPath: string; pageCount: number }> {
  const ext = path.extname(inputPath).toLowerCase().replace('.', '')
  const outputPath = path.join(outputDir, `${outputName}.pdf`)

  if (ext === 'pdf') {
    fs.copyFileSync(inputPath, outputPath)
    const pageCount = await countPdfPages(outputPath)
    return { pdfPath: outputPath, pageCount }
  }

  if (['png', 'jpg', 'jpeg', 'gif', 'webp', 'tiff', 'bmp'].includes(ext)) {
    await convertImageToPdf(inputPath, outputPath)
    return { pdfPath: outputPath, pageCount: 1 }
  }

  if (ext === 'eml') {
    await convertEmlToPdf(inputPath, outputPath)
    const pageCount = await countPdfPages(outputPath)
    return { pdfPath: outputPath, pageCount }
  }

  if (['doc', 'docx', 'odt', 'rtf', 'txt', 'xlsx', 'xls', 'ods', 'pptx', 'ppt', 'odp'].includes(ext)) {
    await convertWithLibreOffice(inputPath, outputDir, outputName)
    const pageCount = await countPdfPages(outputPath)
    return { pdfPath: outputPath, pageCount }
  }

  throw new Error(`Unsupported file format: .${ext}`)
}

async function convertWithLibreOffice(
  inputPath: string,
  outputDir: string,
  outputName: string
): Promise<void> {
  const sofficePath = await findLibreOffice()
  if (!sofficePath) {
    throw new Error(
      'LibreOffice is not installed. Please install LibreOffice to convert Office documents.'
    )
  }

  // LibreOffice names the output based on the input filename
  const inputBaseName = path.basename(inputPath, path.extname(inputPath))
  const tempOutputPath = path.join(outputDir, `${inputBaseName}.pdf`)
  const finalOutputPath = path.join(outputDir, `${outputName}.pdf`)

  await execFileAsync(sofficePath, [
    '--headless',
    '--convert-to',
    'pdf',
    '--outdir',
    outputDir,
    inputPath
  ])

  // Rename to our desired name if different
  if (tempOutputPath !== finalOutputPath && fs.existsSync(tempOutputPath)) {
    fs.renameSync(tempOutputPath, finalOutputPath)
  }
}

async function convertImageToPdf(inputPath: string, outputPath: string): Promise<void> {
  const pdfDoc = await PDFDocument.create()
  const ext = path.extname(inputPath).toLowerCase()

  // Use sharp to convert to PNG buffer for embedding
  const pngBuffer = await sharp(inputPath).png().toBuffer()
  const metadata = await sharp(inputPath).metadata()
  const width = metadata.width ?? 595
  const height = metadata.height ?? 842

  const page = pdfDoc.addPage([width, height])
  const pngImage = await pdfDoc.embedPng(pngBuffer)
  page.drawImage(pngImage, { x: 0, y: 0, width, height })

  const pdfBytes = await pdfDoc.save()
  fs.writeFileSync(outputPath, pdfBytes)
}

async function convertEmlToPdf(inputPath: string, outputPath: string): Promise<void> {
  const raw = fs.readFileSync(inputPath)
  const parsed = await simpleParser(raw)

  const html = parsed.html || `<pre>${parsed.text || ''}</pre>`
  const fullHtml = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        body { font-family: Arial, sans-serif; font-size: 12px; margin: 20px; }
        .header { border-bottom: 1px solid #ccc; margin-bottom: 16px; padding-bottom: 10px; }
        .header div { margin: 4px 0; }
        .label { font-weight: bold; width: 80px; display: inline-block; }
      </style>
    </head>
    <body>
      <div class="header">
        <div><span class="label">From:</span> ${parsed.from?.text ?? ''}</div>
        <div><span class="label">To:</span> ${parsed.to ? (Array.isArray(parsed.to) ? parsed.to.map(a => a.text).join(', ') : parsed.to.text) : ''}</div>
        ${parsed.cc ? `<div><span class="label">CC:</span> ${Array.isArray(parsed.cc) ? parsed.cc.map(a => a.text).join(', ') : parsed.cc.text}</div>` : ''}
        <div><span class="label">Date:</span> ${parsed.date?.toLocaleString() ?? ''}</div>
        <div><span class="label">Subject:</span> ${parsed.subject ?? ''}</div>
      </div>
      <div class="body">${html}</div>
    </body>
    </html>`

  // Use a hidden BrowserWindow to print to PDF
  const win = new BrowserWindow({ show: false, webPreferences: { javascript: true } })
  await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(fullHtml)}`)
  const pdfData = await win.webContents.printToPDF({
    pageSize: 'A4',
    printBackground: true
  })
  win.destroy()
  fs.writeFileSync(outputPath, pdfData)
}

export async function countPdfPages(pdfPath: string): Promise<number> {
  try {
    const bytes = fs.readFileSync(pdfPath)

    // Primary: pdf-lib page-tree traversal.
    // For encrypted PDFs, load() may succeed but getPageCount() silently
    // returns 0 without throwing. Treat count === 0 the same as a throw.
    try {
      const doc = await PDFDocument.load(bytes, { ignoreEncryption: true })
      const count = doc.getPageCount()
      if (count > 0) return count
    } catch { /* fall through */ }

    // Raw-byte fallback: /Count in the Pages dictionary is not encrypted
    // even in owner-protected PDFs. The largest value is the root page count.
    const raw = bytes.toString('latin1')
    const counts = [...raw.matchAll(/\/Count\s+(\d+)/g)].map(m => parseInt(m[1], 10))
    if (counts.length > 0) return Math.max(...counts)

    // Last resort: count individual /Type /Page entries
    const pages = raw.match(/\/Type\s*\/Page[^s]/g)
    return pages ? pages.length : 0
  } catch {
    return 0
  }
}
