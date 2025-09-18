import {heicTo, isHeic} from '../dist/heic-to.js'

const filesInput = document.getElementById("filesInput")
const folderInput = document.getElementById("folderInput")
const excelInput = document.getElementById("excelInput")
const resetButton = document.getElementById("resetButton")
const statusMessage = document.getElementById("statusMessage")
const resultContainer = document.getElementById("resultContainer")
const originalImage = document.getElementById("originalImage")
const convertedImage = document.getElementById("convertedImage")
const progressContainer = document.getElementById("progressContainer")
const progressBar = document.getElementById("progressBar")
const progressText = document.getElementById("progressText")

let excelMapping = new Map()

// Clear all inputs on page load to ensure clean state
function initializePage() {
    filesInput.value = ''
    folderInput.value = ''
    excelInput.value = ''
    excelMapping.clear()

    hideStatus()
    hideResult()
    hideProgress()

    // Reset progress bar
    progressBar.style.width = '0%'
    progressText.textContent = 'Processing...'

    // Clear result containers
    originalImage.innerHTML = ''
    convertedImage.innerHTML = ''

    console.log('🔄 Page initialized with clean state')
}

// Initialize page on load
document.addEventListener('DOMContentLoaded', initializePage)

async function loadJSZip() {
    if (window.JSZip) {
        return window.JSZip
    }

    return new Promise((resolve, reject) => {
        const script = document.createElement('script')
        script.src = './jszip.min.js'
        script.onload = () => {
            if (window.JSZip) {
                resolve(window.JSZip)
            } else {
                reject(new Error('JSZip not loaded'))
            }
        }
        script.onerror = () => reject(new Error('Failed to load JSZip'))
        document.head.appendChild(script)
    })
}

function showStatus(message, type = 'info') {
    statusMessage.textContent = message
    statusMessage.className = `status-message ${type}`
    statusMessage.style.display = 'block'
}

function hideStatus() {
    statusMessage.style.display = 'none'
}

function showResult() {
    resultContainer.style.display = 'grid'
}

function hideResult() {
    resultContainer.style.display = 'none'
}

function showProgress() {
    progressContainer.style.display = 'block'
}

function hideProgress() {
    progressContainer.style.display = 'none'
}

function updateProgress(current, total, message = '') {
    const percentage = Math.round((current / total) * 100)
    progressBar.style.width = `${percentage}%`
    progressText.textContent = message || `Processing ${current}/${total} files...`
}

function createImageElement(src, alt = '') {
    const img = document.createElement('img')
    img.src = src
    img.alt = alt
    img.style.maxWidth = '100%'
    img.style.maxHeight = '300px'
    img.style.objectFit = 'contain'
    return img
}

function clearContainer(container) {
    container.innerHTML = '<p style="color: #666;">Processing...</p>'
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes'
    const k = 1024
    const sizes = ['Bytes', 'KB', 'MB', 'GB']
    const i = Math.floor(Math.log(bytes) / Math.log(k))
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i]
}

async function parseExcelFile(file) {
    try {
        showStatus('📊 Processing Excel file...', 'loading')

        const JSZip = await loadJSZip()

        const zip = new JSZip()
        const zipContent = await zip.loadAsync(file)

        // Get the worksheet data (assuming Sheet1 or first sheet)
        const sheetFile = zipContent.file('xl/worksheets/sheet1.xml') ||
            Object.keys(zipContent.files).find(name => name.includes('worksheets/sheet'))

        if (!sheetFile) {
            throw new Error('Could not find worksheet in Excel file')
        }

        const sheetXml = await sheetFile.async('text')

        // Get shared strings if they exist
        const sharedStringsFile = zipContent.file('xl/sharedStrings.xml')
        let sharedStrings = []
        if (sharedStringsFile) {
            const sharedStringsXml = await sharedStringsFile.async('text')
            const parser = new DOMParser()
            const sharedStringsDoc = parser.parseFromString(sharedStringsXml, 'text/xml')
            const stringNodes = sharedStringsDoc.getElementsByTagName('t')
            sharedStrings = Array.from(stringNodes).map(node => node.textContent)
        }

        // Parse the worksheet XML
        const parser = new DOMParser()
        const doc = parser.parseFromString(sheetXml, 'text/xml')
        const rows = doc.getElementsByTagName('row')

        const mapping = new Map()

        // Process rows to extract column Q (17) and R (18) data
        for (const row of rows) {
            const cells = row.getElementsByTagName('c')
            let columnQValue = ''
            let columnRValue = ''

            for (const cell of cells) {
                const cellRef = cell.getAttribute('r')
                const column = cellRef.match(/[A-Z]+/)[0]

                if (column === 'Q' || column === 'R') {
                    const valueElement = cell.getElementsByTagName('v')[0]
                    if (valueElement) {
                        let value = valueElement.textContent

                        // Check if it's a shared string reference
                        const type = cell.getAttribute('t')
                        if (type === 's' && sharedStrings.length > 0) {
                            const index = parseInt(value)
                            value = sharedStrings[index] || value
                        }

                        if (column === 'Q') {
                            columnQValue = value
                        } else if (column === 'R') {
                            columnRValue = value
                        }
                    }
                }
            }

            // Add mapping if both values exist
            if (columnQValue && columnRValue) {
                mapping.set(columnQValue, columnRValue)
            }
        }

        if (mapping.size === 0) {
            showStatus('⚠️ No file mappings found in Excel columns Q and R', 'warning')
        } else {
            showStatus(`✅ Loaded ${mapping.size} file mappings from Excel`, 'success')
        }

        return mapping

    } catch (error) {
        console.error('Excel parsing error:', error)
        showStatus(`❌ Failed to parse Excel file: ${error.message}`, 'error')
        return new Map()
    }
}

async function parseCSVFile(file) {
    try {
        const text = await file.text()
        const lines = text.split('\n')
        const mapping = new Map()

        for (const line of lines) {
            const columns = line.split(',').map(col => col.trim().replace(/"/g, ''))

            // Assuming columns Q and R are at indices 16 and 17 (0-based)
            if (columns.length >= 18) {
                const columnQ = columns[16]
                const columnR = columns[17]

                if (columnQ && columnR) {
                    mapping.set(columnQ, columnR)
                }
            }
        }

        return mapping

    } catch (error) {
        console.error('CSV parsing error:', error)
        showStatus(`❌ Failed to parse CSV file: ${error.message}`, 'error')
        return new Map()
    }
}

function getNewFileName(originalFileName, mapping) {
    const newName = mapping.get(originalFileName)
    return newName || originalFileName
}

async function handleFiles(files) {
    if (!files || files.length === 0) {
        hideStatus()
        hideResult()
        return
    }

    if (files.length === 1) {
        await handleSingleFile(files[0])
    } else {
        const fileItems = files.map(file => ({
            file,
            path: file.name,
            relativePath: file.webkitRelativePath ? file.webkitRelativePath.replace(file.name, '') : ''
        }))
        await processBatchConversion(fileItems)
    }
}

async function processBatchConversion(fileItems) {
    try {
        hideResult()
        showProgress()

        const JSZip = await loadJSZip()

        const heicFiles = []
        const nonHeicFiles = []

        updateProgress(0, fileItems.length, 'Analyzing files...')

        for (let i = 0; i < fileItems.length; i++) {
            const fileItem = fileItems[i]
            const isHeicFile = await isHeic(fileItem.file)

            // Apply Excel mapping to rename files before processing
            const originalFileName = fileItem.file.name
            const newFileName = getNewFileName(originalFileName, excelMapping)

            // Create a new file object with the renamed filename if mapping exists
            let processedFile = fileItem.file
            if (newFileName !== originalFileName) {
                // Create new file with renamed filename
                processedFile = new File([fileItem.file], newFileName, {
                    type: fileItem.file.type,
                    lastModified: fileItem.file.lastModified
                })

                // Update the fileItem
                fileItem.file = processedFile
                fileItem.originalName = originalFileName
                fileItem.renamedFrom = originalFileName
            }

            if (isHeicFile) {
                heicFiles.push(fileItem)
            } else {
                nonHeicFiles.push(fileItem)
            }

            updateProgress(i + 1, fileItems.length, 'Analyzing files...')
        }

        if (heicFiles.length === 0 && nonHeicFiles.length === 0) {
            showStatus(`ℹ️ No image files found.`, 'info')
            hideProgress()
            return
        }

        showStatus(`🔄 Converting ${heicFiles.length} HEIC files...`, 'loading')

        const zip = new JSZip()
        const convertedFiles = []
        const failedFiles = []
        const renamedFiles = []

        for (let i = 0; i < heicFiles.length; i++) {
            const fileItem = heicFiles[i]
            const fileName = fileItem.file.name

            try {
                updateProgress(i + 1, heicFiles.length, `Converting ${fileName}...`)

                const convertedBlob = await heicTo({
                    blob: fileItem.file,
                    type: 'image/jpeg',
                    quality: 0.8
                })

                const outputFileName = fileName.replace(/\.(heic|heif)$/i, '.jpg')
                const zipPath = fileItem.relativePath + outputFileName

                zip.file(zipPath, convertedBlob)
                convertedFiles.push({
                    original: fileItem.renamedFrom || fileName,
                    converted: outputFileName,
                    originalSize: fileItem.file.size,
                    convertedSize: convertedBlob.size,
                    wasRenamed: !!fileItem.renamedFrom
                })

                if (fileItem.renamedFrom) {
                    renamedFiles.push({
                        from: fileItem.renamedFrom,
                        to: fileName
                    })
                }

            } catch (error) {
                console.error(`Failed to convert ${fileName}:`, error)
                failedFiles.push({
                    name: fileName,
                    error: error.message
                })
            }
        }

        // Add non-HEIC image files to ZIP as-is
        for (const fileItem of nonHeicFiles) {
            const zipPath = fileItem.relativePath + fileItem.file.name
            zip.file(zipPath, fileItem.file)

            if (fileItem.renamedFrom) {
                renamedFiles.push({
                    from: fileItem.renamedFrom,
                    to: fileItem.file.name
                })
            }
        }

        if (convertedFiles.length === 0 && nonHeicFiles.length === 0) {
            showStatus('❌ No files could be processed', 'error')
            hideProgress()
            return
        }

        updateProgress(1, 1, 'Creating ZIP file...')

        const zipBlob = await zip.generateAsync({type: 'blob'})

        hideProgress()
        showBatchResults(convertedFiles, failedFiles, nonHeicFiles.length, zipBlob, renamedFiles)

    } catch (error) {
        console.error('Batch conversion error:', error)
        showStatus(`❌ Error: ${error.message || 'Failed to process files'}`, 'error')
        hideProgress()
    }
}

function showBatchResults(convertedFiles, failedFiles, otherFiles, zipBlob, renamedFiles = []) {
    originalImage.innerHTML = `
    <div style="padding: 20px; text-align: center;">
      <div style="font-size: 3rem; margin-bottom: 10px;">📁</div>
      <div style="font-weight: bold; margin-bottom: 10px;">Batch Conversion Results</div>
      <div style="color: #666; font-size: 0.9rem; line-height: 1.5;">
        <div>✅ ${convertedFiles.length} HEIC files converted</div>
        ${renamedFiles.length > 0 ? `<div>🏷️ ${renamedFiles.length} files renamed</div>` : ''}
        ${failedFiles.length > 0 ? `<div style="color: #dc3545;">❌ ${failedFiles.length} files failed</div>` : ''}
      </div>
    </div>
  `

    const zipUrl = URL.createObjectURL(zipBlob)
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '')
    const zipFileName = `converted_images_${timestamp}.zip`

    convertedImage.innerHTML = `
    <div style="padding: 20px; text-align: center;">
      <div style="font-size: 3rem; margin-bottom: 10px;">📦</div>
      <div style="font-weight: bold; margin-bottom: 10px;">Download Ready</div>
      <div style="color: #666; font-size: 0.9rem; margin-bottom: 15px;">
        ZIP • ${formatFileSize(zipBlob.size)}
      </div>
      <button onclick="downloadZip('${zipUrl}', '${zipFileName}')" 
              style="background: #28a745; color: white; border: none; padding: 12px 24px; border-radius: 4px; cursor: pointer; font-size: 1rem; margin-bottom: 15px;">
        Download ZIP
      </button>
      ${renamedFiles.length > 0 ? `
        <details style="margin-bottom: 15px; text-align: left; max-width: 300px; margin-left: auto; margin-right: auto;">
          <summary style="cursor: pointer; color: #28a745;">Renamed Files (${renamedFiles.length})</summary>
          <div style="font-size: 0.8rem; margin-top: 5px; padding-left: 15px;">
            ${renamedFiles.map(f => `<div>• ${f.from} → ${f.to}</div>`).join('')}
          </div>
        </details>
      ` : ''}
      ${failedFiles.length > 0 ? `
        <details style="margin-top: 15px; text-align: left; max-width: 300px; margin-left: auto; margin-right: auto;">
          <summary style="cursor: pointer; color: #dc3545;">Failed Files (${failedFiles.length})</summary>
          <div style="font-size: 0.8rem; margin-top: 5px; padding-left: 15px;">
            ${failedFiles.map(f => `<div>• ${f.name}: ${f.error}</div>`).join('')}
          </div>
        </details>
      ` : ''}
    </div>
  `

    showResult()

    const successCount = convertedFiles.length + otherFiles
    const totalFiles = successCount + failedFiles.length

    if (failedFiles.length > 0) {
        showStatus(
            `⚠️ Processed ${successCount}/${totalFiles} files. ${failedFiles.length} files failed to convert.`,
            'warning'
        )
    } else {
        let message = `🎉 Successfully processed all ${totalFiles} files!`
        if (renamedFiles.length > 0) {
            message += ` ${renamedFiles.length} files were renamed using Excel mappings.`
        }
        showStatus(message, 'success')
    }
}

async function handleSingleFile(file) {
    try {
        hideResult()
        hideProgress()
        clearContainer(originalImage)
        clearContainer(convertedImage)

        showStatus('🔍 Analyzing file...', 'loading')

        const originalFileName = file.name
        const newFileName = getNewFileName(originalFileName, excelMapping)
        let processedFile = file
        let wasRenamed = false

        if (newFileName !== originalFileName) {
            processedFile = new File([file], newFileName, {
                type: file.type,
                lastModified: file.lastModified
            })
            wasRenamed = true
        }

        const isHeicFile = await isHeic(processedFile)

        if (isHeicFile) {
            showStatus('✅ HEIC/HEIF file detected! Converting to JPEG...', 'loading')

            originalImage.innerHTML = `
        <div style="padding: 20px; text-align: center;">
          <div style="font-size: 3rem; margin-bottom: 10px;">📷</div>
          <div style="font-weight: bold; margin-bottom: 5px;">${processedFile.name}</div>
          ${wasRenamed ? `<div style="color: #28a745; font-size: 0.8rem; margin-bottom: 5px;">Renamed from: ${originalFileName}</div>` : ''}
          <div style="color: #666; font-size: 0.9rem;">
            HEIC/HEIF • ${formatFileSize(processedFile.size)}
          </div>
        </div>
      `

            showResult()

            const startTime = Date.now()
            const convertedBlob = await heicTo({
                blob: processedFile,
                type: 'image/jpeg',
                quality: 0.8
            })
            const conversionTime = Date.now() - startTime

            const convertedUrl = URL.createObjectURL(convertedBlob)
            const convertedImg = createImageElement(convertedUrl, 'Converted JPEG image')

            convertedImage.innerHTML = ''
            convertedImage.appendChild(convertedImg)

            const infoDiv = document.createElement('div')
            infoDiv.style.marginTop = '15px'
            infoDiv.style.textAlign = 'center'
            const downloadFileName = processedFile.name.replace(/\.[^/.]+$/, '') + '.jpg'
            infoDiv.innerHTML = `
        <div style="color: #666; font-size: 0.9rem; margin-bottom: 10px;">
          JPEG • ${formatFileSize(convertedBlob.size)} • ${conversionTime}ms
        </div>
        <button onclick="downloadImage('${convertedUrl}', '${downloadFileName}')" 
                style="background: #28a745; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; font-size: 0.9rem;">
          📥 Download JPEG
        </button>
      `
            convertedImage.appendChild(infoDiv)

            const compressionRatio = ((processedFile.size - convertedBlob.size) / processedFile.size * 100).toFixed(1)
            const sizeChange = convertedBlob.size > processedFile.size ? 'larger' : 'smaller'

            let statusMsg = `🎉 Successfully converted! File is ${Math.abs(compressionRatio)}% ${sizeChange} (${conversionTime}ms)`
            if (wasRenamed) {
                statusMsg += ` File was renamed using Excel mapping.`
            }

            showStatus(statusMsg, 'success')

        } else {
            // Not a HEIC file - show as regular image if possible
            showStatus('ℹ️ This is not a HEIC/HEIF file. Displaying as regular image.', 'info')

            if (processedFile.type.startsWith('image/')) {
                const imageUrl = URL.createObjectURL(processedFile)
                const img = createImageElement(imageUrl, 'Original image')

                originalImage.innerHTML = ''
                originalImage.appendChild(img)

                const infoDiv = document.createElement('div')
                infoDiv.style.marginTop = '15px'
                infoDiv.style.textAlign = 'center'
                infoDiv.innerHTML = `
          <div style="font-weight: bold; margin-bottom: 5px;">${processedFile.name}</div>
          ${wasRenamed ? `<div style="color: #28a745; font-size: 0.8rem; margin-bottom: 5px;">Renamed from: ${originalFileName}</div>` : ''}
          <div style="color: #666; font-size: 0.9rem;">
            ${processedFile.type} • ${formatFileSize(processedFile.size)}
          </div>
        `
                originalImage.appendChild(infoDiv)

                convertedImage.innerHTML = `
          <div style="padding: 20px; text-align: center; color: #666;">
            <div style="font-size: 2rem; margin-bottom: 10px;">ℹ️</div>
            <div>No conversion needed</div>
            <div style="font-size: 0.9rem; margin-top: 5px;">
              This file is already in a web-compatible format
            </div>
            ${wasRenamed ? `<div style="color: #28a745; font-size: 0.8rem; margin-top: 5px;">File was renamed using Excel mapping</div>` : ''}
          </div>
        `

                showResult()
            } else {
                originalImage.innerHTML = `
          <div style="padding: 20px; text-align: center; color: #666;">
            <div style="font-size: 2rem; margin-bottom: 10px;">❌</div>
            <div>Unsupported file type</div>
            <div style="font-size: 0.9rem; margin-top: 5px;">
              Please select a HEIC/HEIF or image file
            </div>
          </div>
        `
                showResult()
            }
        }

    } catch (error) {
        console.error('Conversion error:', error)
        showStatus(`❌ Error: ${error.message || 'Failed to process file'}`, 'error')
        hideResult()
    }
}

excelInput.addEventListener('change', async (event) => {
    const file = event.target.files[0]

    if (!file) {
        excelMapping.clear()
        return
    }

    const fileName = file.name.toLowerCase()

    if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        excelMapping = await parseExcelFile(file)
    } else if (fileName.endsWith('.csv')) {
        excelMapping = await parseCSVFile(file)
    } else {
        showStatus('❌ Please select an Excel (.xlsx, .xls) or CSV file', 'error')
        event.target.value = '' // Clear the input
        return
    }
})

filesInput.addEventListener('change', async (event) => {
    const files = Array.from(event.target.files)
    folderInput.value = ''
    await handleFiles(files)
})

folderInput.addEventListener('change', async (event) => {
    const files = Array.from(event.target.files)
    filesInput.value = ''
    await handleFiles(files)
})

window.downloadImage = function (url, filename) {
    const a = document.createElement('a')
    a.href = url
    a.download = filename
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
}

window.downloadZip = function (url, filename) {
    const a = document.createElement('a')
    a.href = url
    a.download = filename
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
}

window.addEventListener('beforeunload', () => {
    const images = document.querySelectorAll('img[src^="blob:"]')
    images.forEach(img => URL.revokeObjectURL(img.src))
})