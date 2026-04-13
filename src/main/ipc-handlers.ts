import { ipcMain, dialog, shell, BrowserWindow, app } from 'electron'
import path from 'path'
import fs from 'fs'
import os from 'os'
import { v4 as uuidv4 } from 'uuid'
import {
  listProjects,
  getProject,
  createProject,
  updateProject,
  deleteProject,
  listExhibits,
  getExhibit,
  createExhibit,
  updateExhibit,
  deleteExhibits,
  shiftOrderIndexes,
  getCitationsForProject,
  createCitation,
  getExhibitNotes,
  setExhibitNotes
} from './database'
import { convertToPdf, findLibreOffice, countPdfPages } from './converter'
import { citeExhibit, refreshCitations, listWordDocuments } from './word-integration'
import { exportBundle } from './exporter'
import { Project, Exhibit, ExportOptions } from '../shared/types'

const DEFAULT_ROOT = path.join(os.homedir(), 'ExhibitManager')

export function registerIpcHandlers(mainWindow: BrowserWindow): void {
  // ── System ──────────────────────────────────────────────────────────────────

  ipcMain.handle('system:selectFiles', async (_e, filters: Electron.FileFilter[]) => {
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openFile', 'multiSelections'],
      filters
    })
    return result.canceled ? [] : result.filePaths
  })

  ipcMain.handle('system:selectDirectory', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openDirectory']
    })
    return result.canceled ? null : result.filePaths[0]
  })

  ipcMain.handle('system:openFile', (_e, filePath: string) => {
    shell.openPath(filePath)
  })

  ipcMain.handle('system:checkLibreOffice', async () => {
    const found = await findLibreOffice()
    return { installed: !!found, path: found }
  })

  ipcMain.handle('system:getResourcesPath', () => app.isPackaged
    ? path.join(process.resourcesPath)
    : path.join(app.getAppPath(), 'resources')
  )

  // ── Projects ────────────────────────────────────────────────────────────────

  ipcMain.handle('project:list', () => listProjects())

  ipcMain.handle('project:create', (_e, name: string, prefix: string, separator: string, citationTemplate?: string) => {
    const id = uuidv4()
    const rootPath = path.join(DEFAULT_ROOT, id)
    fs.mkdirSync(path.join(rootPath, 'originals'), { recursive: true })
    fs.mkdirSync(path.join(rootPath, 'pdfs'), { recursive: true })
    const now = new Date().toISOString()
    const project: Project = {
      id,
      name,
      prefix,
      separator,
      citationTemplate: citationTemplate ?? 'Exhibit {label}, {description}, {date}',
      rootPath,
      createdAt: now,
      updatedAt: now
    }
    return createProject(project)
  })

  ipcMain.handle('project:open', (_e, id: string) => getProject(id))

  ipcMain.handle('project:update', (_e, project: Project) => updateProject(project))

  ipcMain.handle('project:delete', (_e, id: string) => {
    const project = getProject(id)
    if (project) {
      try { fs.rmSync(project.rootPath, { recursive: true, force: true }) } catch { /* ignore */ }
    }
    deleteProject(id)
  })

  // ── Exhibits ─────────────────────────────────────────────────────────────────

  ipcMain.handle('exhibit:list', async (_e, projectId: string) => {
    const exhibits = listExhibits(projectId)
    // Background-fix page counts that were stored as 0 (e.g. imported before
    // the encrypted-PDF fix landed). Re-run countPdfPages for any zero-count
    // exhibit whose PDF file still exists, then persist the corrected value.
    for (const ex of exhibits) {
      if (ex.pageCount === 0 && fs.existsSync(ex.pdfPath)) {
        try {
          const n = await countPdfPages(ex.pdfPath)
          if (n > 0) {
            updateExhibit(ex.id, { pageCount: n })
            ex.pageCount = n
          }
        } catch { /* leave as 0 if still unreadable */ }
      }
    }
    return exhibits
  })

  ipcMain.handle(
    'exhibit:add',
    async (_e, projectId: string, filePaths: string[], insertAfterIndex?: number) => {
      const project = getProject(projectId)
      if (!project) throw new Error('Project not found')

      const existing = listExhibits(projectId)
      let startIndex = insertAfterIndex != null ? insertAfterIndex + 1 : existing.length + 1

      // Shift existing exhibits if inserting in the middle
      if (insertAfterIndex != null && insertAfterIndex < existing.length) {
        shiftOrderIndexes(projectId, startIndex, filePaths.length)
        // Rename existing PDFs that were shifted
        const shifted = existing.filter(e => e.orderIndex >= startIndex)
        for (const ex of shifted.reverse()) {
          const newIndex = ex.orderIndex + filePaths.length
          const newLabel = `${project.prefix}${project.separator}${newIndex}`
          const oldPdfPath = ex.pdfPath
          const newPdfPath = path.join(project.rootPath, 'pdfs', `${newLabel}.pdf`)
          if (fs.existsSync(oldPdfPath)) fs.renameSync(oldPdfPath, newPdfPath)
          updateExhibit(ex.id, { orderIndex: newIndex, label: newLabel, pdfPath: newPdfPath })
        }
      }

      const results: Exhibit[] = []
      for (let i = 0; i < filePaths.length; i++) {
        const filePath = filePaths[i]
        const orderIndex = startIndex + i
        const label = `${project.prefix}${project.separator}${orderIndex}`
        const originalFilename = path.basename(filePath)
        const originalFormat = path.extname(filePath).toLowerCase().replace('.', '')

        // Notify progress
        mainWindow.webContents.send('exhibit:convertProgress', {
          filename: originalFilename,
          status: 'converting'
        })

        try {
          // Copy original
          const origDest = path.join(project.rootPath, 'originals', originalFilename)
          fs.copyFileSync(filePath, origDest)

          // Convert to PDF
          const pdfsDir = path.join(project.rootPath, 'pdfs')
          const { pdfPath, pageCount } = await convertToPdf(filePath, pdfsDir, label)

          const exhibit: Exhibit = {
            id: uuidv4(),
            projectId,
            orderIndex,
            label,
            originalFilename,
            originalFormat,
            description: '',
            dateOfDocument: null,
            pdfPath,
            pageCount,
            createdAt: new Date().toISOString()
          }
          createExhibit(exhibit)
          results.push(exhibit)

          mainWindow.webContents.send('exhibit:convertProgress', {
            filename: originalFilename,
            status: 'done'
          })
        } catch (err) {
          mainWindow.webContents.send('exhibit:convertProgress', {
            filename: originalFilename,
            status: 'error',
            error: String(err)
          })
        }
      }

      return results
    }
  )

  ipcMain.handle(
    'exhibit:reorder',
    (_e, projectId: string, exhibitId: string, newIndex: number) => {
      const exhibits = listExhibits(projectId)
      const project = getProject(projectId)
      if (!project) throw new Error('Project not found')

      const exhibit = exhibits.find(e => e.id === exhibitId)
      if (!exhibit) throw new Error('Exhibit not found')

      const oldIndex = exhibit.orderIndex
      if (oldIndex === newIndex) return exhibits

      // Reorder: remove from old position, insert at new
      const reordered = exhibits.filter(e => e.id !== exhibitId)
      reordered.splice(newIndex - 1, 0, exhibit)

      // Reassign orderIndexes and rename PDFs
      for (let i = 0; i < reordered.length; i++) {
        const ex = reordered[i]
        const newOrderIndex = i + 1
        if (ex.orderIndex !== newOrderIndex) {
          const newLabel = `${project.prefix}${project.separator}${newOrderIndex}`
          const newPdfPath = path.join(project.rootPath, 'pdfs', `${newLabel}.pdf`)
          if (fs.existsSync(ex.pdfPath) && ex.pdfPath !== newPdfPath) {
            // Use temp rename to avoid conflicts
            fs.renameSync(ex.pdfPath, newPdfPath + '.tmp')
          }
          updateExhibit(ex.id, { orderIndex: newOrderIndex, label: newLabel, pdfPath: newPdfPath })
        }
      }
      // Remove .tmp extensions
      for (let i = 0; i < reordered.length; i++) {
        const newLabel = `${project.prefix}${project.separator}${i + 1}`
        const tmpPath = path.join(project.rootPath, 'pdfs', `${newLabel}.pdf.tmp`)
        const finalPath = path.join(project.rootPath, 'pdfs', `${newLabel}.pdf`)
        if (fs.existsSync(tmpPath)) fs.renameSync(tmpPath, finalPath)
      }

      return listExhibits(projectId)
    }
  )

  ipcMain.handle(
    'exhibit:update',
    (_e, exhibitId: string, fields: Partial<Pick<Exhibit, 'description' | 'dateOfDocument'>>) => {
      return updateExhibit(exhibitId, fields)
    }
  )

  ipcMain.handle('exhibit:delete', (_e, projectId: string, exhibitIds: string[]) => {
    const project = getProject(projectId)
    if (!project) throw new Error('Project not found')

    for (const id of exhibitIds) {
      const ex = getExhibit(id)
      if (ex) {
        try { fs.unlinkSync(ex.pdfPath) } catch { /* ignore */ }
        // Keep original file
      }
    }
    deleteExhibits(exhibitIds)

    // Renumber remaining exhibits
    const remaining = listExhibits(projectId)
    for (let i = 0; i < remaining.length; i++) {
      const ex = remaining[i]
      const newOrderIndex = i + 1
      if (ex.orderIndex !== newOrderIndex) {
        const newLabel = `${project.prefix}${project.separator}${newOrderIndex}`
        const newPdfPath = path.join(project.rootPath, 'pdfs', `${newLabel}.pdf`)
        if (fs.existsSync(ex.pdfPath) && ex.pdfPath !== newPdfPath) {
          fs.renameSync(ex.pdfPath, newPdfPath)
        }
        updateExhibit(ex.id, { orderIndex: newOrderIndex, label: newLabel, pdfPath: newPdfPath })
      }
    }

    return listExhibits(projectId)
  })

  ipcMain.handle('exhibit:getPdf', (_e, exhibitId: string) => {
    const ex = getExhibit(exhibitId)
    return ex?.pdfPath ?? null
  })

  ipcMain.handle('fs:readFileBytes', (_e, filePath: string) => {
    return fs.readFileSync(filePath)  // returns Buffer, serialised as Uint8Array over IPC
  })

  ipcMain.handle('exhibit:getNotes', (_e, exhibitId: string) => getExhibitNotes(exhibitId))
  ipcMain.handle('exhibit:setNotes', (_e, exhibitId: string, notes: string) => { setExhibitNotes(exhibitId, notes) })

  // ── Word Integration ─────────────────────────────────────────────────────────

  ipcMain.handle('word:listDocs', async () => {
    return listWordDocuments()
  })

  ipcMain.handle('word:cite', async (_e, projectId: string, exhibitId: string, docPath?: string) => {
    const project = getProject(projectId)
    const exhibit = getExhibit(exhibitId)
    if (!project || !exhibit) return { success: false, error: 'Not found' }

    const result = await citeExhibit(exhibit, project.citationTemplate, docPath)
    if (result.success && result.citation) {
      createCitation(result.citation)
    }
    return result
  })

  ipcMain.handle('word:refresh', async (_e, projectId: string, docPath?: string) => {
    const project = getProject(projectId)
    if (!project) return { updated: 0, orphaned: 0, error: 'Project not found' }
    const exhibits = listExhibits(projectId)
    return refreshCitations(exhibits, project.citationTemplate, docPath)
  })

  // ── Export ───────────────────────────────────────────────────────────────────

  ipcMain.handle('export:bundle', async (_e, projectId: string, options: ExportOptions) => {
    const project = getProject(projectId)
    if (!project) throw new Error('Project not found')
    const exhibits = listExhibits(projectId)

    const { filePath, canceled } = await dialog.showSaveDialog(mainWindow, {
      title: 'Save Exhibit Bundle',
      defaultPath: path.join(os.homedir(), `${project.name} - Exhibit Bundle ${new Date().toISOString().slice(0,10)}.pdf`),
      filters: [{ name: 'PDF', extensions: ['pdf'] }]
    })
    if (canceled || !filePath) return null

    const outputPath = await exportBundle(project, exhibits, options, filePath)
    shell.openPath(outputPath)
    return outputPath
  })
}
