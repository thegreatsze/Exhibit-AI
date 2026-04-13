import { PDFDocument, PDFPage, rgb, StandardFonts, PageSizes } from 'pdf-lib'
import fs from 'fs'
import { Exhibit, Project, ExportOptions } from '../shared/types'

export async function exportBundle(
  project: Project,
  exhibits: Exhibit[],
  options: ExportOptions,
  outputPath: string
): Promise<string> {
  const mergedPdf = await PDFDocument.create()
  const font     = await mergedPdf.embedFont(StandardFonts.Helvetica)
  const boldFont = await mergedPdf.embedFont(StandardFonts.HelveticaBold)
  const pageSize = options.pageSize === 'Letter' ? PageSizes.Letter : PageSizes.A4
  const [pageWidth, pageHeight] = pageSize
  const margin = 60

  // Running 1-indexed page counter. Cover = 1 and is never stamped.
  let pageNum = options.includeCoverPage ? 2 : 1

  function stampNumber(page: PDFPage, w: number, h: number): void {
    if (!options.includePageNumbers) { pageNum++; return }
    const numStr = String(pageNum)
    const numWidth = font.widthOfTextAtSize(numStr, 9)
    const isOdd = pageNum % 2 !== 0
    let x: number, y: number
    switch (options.pageNumberPosition) {
      case 'top-outer':   x = isOdd ? margin : w - margin - numWidth; y = h - 28; break
      case 'top-left':    x = margin;                                  y = h - 28; break
      case 'top-right':   x = w - margin - numWidth;                  y = h - 28; break
      case 'top-center':  x = (w - numWidth) / 2;                     y = h - 28; break
      case 'bottom-center': x = (w - numWidth) / 2;                   y = 20;     break
      default:            x = margin;                                  y = h - 28;
    }
    page.drawText(numStr, { x, y, size: 9, font, color: rgb(0.45, 0.45, 0.45) })
    pageNum++
  }

  // ── Pre-load source PDFs ──────────────────────────────────────────────────────
  const srcDocs: (PDFDocument | null)[] = []
  const validExhibits: Exhibit[] = []
  for (const exhibit of exhibits) {
    if (!fs.existsSync(exhibit.pdfPath)) continue
    try {
      srcDocs.push(await PDFDocument.load(fs.readFileSync(exhibit.pdfPath), { ignoreEncryption: true }))
    } catch {
      srcDocs.push(null)   // mark as unloadable; handled in the render loop
    }
    validExhibits.push(exhibit)
  }

  // ── TOC page-number calculation ───────────────────────────────────────────────
  // Layout: [cover?] [divider_0] [exhibit_0 pages…] [divider_1] [exhibit_1 pages…] …
  const dividerPageNums: number[] = []
  let toc = options.includeCoverPage ? 2 : 1
  for (let i = 0; i < validExhibits.length; i++) {
    dividerPageNums.push(toc)
    let pc = 1
    try { pc = srcDocs[i]?.getPageCount() ?? 1 } catch { pc = 1 }
    toc += 1 + pc
  }

  // ── Cover page (not numbered) ─────────────────────────────────────────────────
  if (options.includeCoverPage) {
    const cover = mergedPdf.addPage(pageSize)
    let y = pageHeight - margin

    cover.drawText(options.coverTitle || project.name, {
      x: margin, y, size: 24, font: boldFont, color: rgb(0.1, 0.14, 0.2)
    })
    y -= 36

    if (options.coverSubtitle) {
      cover.drawText(options.coverSubtitle, { x: margin, y, size: 16, font, color: rgb(0.4, 0.4, 0.4) })
      y -= 20
    }

    if (options.coverShowDate) {
      cover.drawText(`Generated: ${new Date().toLocaleDateString()}`, {
        x: margin, y, size: 10, font, color: rgb(0.5, 0.5, 0.5)
      })
      y -= 16
    }

    y -= 24
    cover.drawLine({ start: { x: margin, y }, end: { x: pageWidth - margin, y }, thickness: 1, color: rgb(0.8, 0.8, 0.8) })
    y -= 24

    if (options.coverShowToc) {
      cover.drawText('TABLE OF CONTENTS', { x: margin, y, size: 12, font: boldFont, color: rgb(0.1, 0.14, 0.2) })
      y -= 24
      cover.drawText('Exhibit',     { x: margin,                   y, size: 9, font: boldFont, color: rgb(0.3, 0.3, 0.3) })
      cover.drawText('Description', { x: margin + 80,              y, size: 9, font: boldFont, color: rgb(0.3, 0.3, 0.3) })
      cover.drawText('Date',        { x: pageWidth - margin - 120, y, size: 9, font: boldFont, color: rgb(0.3, 0.3, 0.3) })
      cover.drawText('Page',        { x: pageWidth - margin - 30,  y, size: 9, font: boldFont, color: rgb(0.3, 0.3, 0.3) })
      y -= 4
      cover.drawLine({ start: { x: margin, y }, end: { x: pageWidth - margin, y }, thickness: 0.5, color: rgb(0.7, 0.7, 0.7) })
      y -= 16
      for (let i = 0; i < validExhibits.length; i++) {
        if (y < 60) break
        const ex = validExhibits[i]
        const desc = ex.description.length > 60 ? ex.description.slice(0, 57) + '...' : ex.description
        cover.drawText(ex.label,                  { x: margin,                   y, size: 9, font, color: rgb(0, 0, 0) })
        cover.drawText(desc,                       { x: margin + 80,              y, size: 9, font, color: rgb(0, 0, 0) })
        cover.drawText(ex.dateOfDocument ?? '',    { x: pageWidth - margin - 120, y, size: 9, font, color: rgb(0, 0, 0) })
        cover.drawText(String(dividerPageNums[i]), { x: pageWidth - margin - 30,  y, size: 9, font, color: rgb(0, 0, 0) })
        y -= 16
      }
    }
    // Cover = page 1: do NOT call stampNumber
  }

  // ── Exhibits: divider + pages ─────────────────────────────────────────────────
  for (let i = 0; i < srcDocs.length; i++) {
    const srcDoc = srcDocs[i]
    const exhibit = validExhibits[i]

    // ── Divider page ─────────────────────────────────────────────────────────
    const divider = mergedPdf.addPage(pageSize)
    const cx = pageWidth / 2, cy = pageHeight / 2

    divider.drawRectangle({
      x: margin, y: cy - 55, width: pageWidth - margin * 2, height: 110,
      color: rgb(0.95, 0.96, 0.98),
      borderColor: rgb(0.75, 0.80, 0.90), borderWidth: 1.5
    })
    const labelW = boldFont.widthOfTextAtSize(exhibit.label, 42)
    divider.drawText(exhibit.label, {
      x: cx - labelW / 2, y: cy + 4, size: 42, font: boldFont, color: rgb(0.1, 0.14, 0.2)
    })
    if (exhibit.description) {
      const descText = exhibit.description.length > 80
        ? exhibit.description.slice(0, 77) + '...' : exhibit.description
      const descW = font.widthOfTextAtSize(descText, 11)
      divider.drawText(descText, { x: cx - descW / 2, y: cy - 30, size: 11, font, color: rgb(0.4, 0.4, 0.4) })
    }
    stampNumber(divider, pageWidth, pageHeight)

    // ── Exhibit pages ─────────────────────────────────────────────────────────
    // Guard the entire page-embedding block. For encrypted or structurally
    // unusual PDFs, pdf-lib may throw at getPageCount(), getPage(), or
    // embedPages() with "Expected instance of PDFDict" errors. Catching here
    // prevents one bad exhibit from aborting the whole export; a placeholder
    // page is inserted instead.
    if (srcDoc === null) {
      // PDF failed to load entirely
      const [w, h] = pageSize
      const errPage = mergedPdf.addPage([w, h])
      const msg = 'This exhibit could not be loaded (encrypted or unsupported PDF).'
      errPage.drawText(msg, { x: margin, y: h / 2, size: 9, font, color: rgb(0.6, 0.2, 0.2) })
      stampNumber(errPage, w, h)
      continue
    }

    let srcPageCount: number
    try { srcPageCount = srcDoc.getPageCount() } catch { srcPageCount = 0 }

    if (srcPageCount === 0) {
      const [w, h] = pageSize
      const errPage = mergedPdf.addPage([w, h])
      const msg = 'This exhibit could not be embedded (encrypted or unsupported PDF).'
      errPage.drawText(msg, { x: margin, y: h / 2, size: 9, font, color: rgb(0.6, 0.2, 0.2) })
      stampNumber(errPage, w, h)
      continue
    }

    let embedded: Awaited<ReturnType<typeof mergedPdf.embedPages>>
    try {
      embedded = await mergedPdf.embedPages(
        Array.from({ length: srcPageCount }, (_, j) => srcDoc.getPage(j))
      )
    } catch {
      const [w, h] = pageSize
      const errPage = mergedPdf.addPage([w, h])
      const msg = 'This exhibit could not be embedded (encrypted or unsupported PDF).'
      errPage.drawText(msg, { x: margin, y: h / 2, size: 9, font, color: rgb(0.6, 0.2, 0.2) })
      stampNumber(errPage, w, h)
      continue
    }

    for (let j = 0; j < srcPageCount; j++) {
      let w: number, h: number
      try {
        const sz = srcDoc.getPage(j).getSize()
        w = sz.width; h = sz.height
      } catch {
        w = pageWidth; h = pageHeight
      }
      const newPage = mergedPdf.addPage([w, h])

      // Paint the original page content
      newPage.drawPage(embedded[j])

      // Optional exhibit label footer stamp
      if (options.includeStamps) {
        const stampW = font.widthOfTextAtSize(exhibit.label, 8)
        newPage.drawText(exhibit.label, {
          x: w / 2 - stampW / 2, y: 12,
          size: 8, font, color: rgb(0.5, 0.5, 0.5)
        })
      }

      stampNumber(newPage, w, h)
    }
  }

  fs.writeFileSync(outputPath, await mergedPdf.save())
  return outputPath
}
