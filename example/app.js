import {heicTo, isHeic} from '../dist/heic-to.js'

const filesInput = document.getElementById("filesInput")
const folderInput = document.getElementById("folderInput")
const excelInput = document.getElementById("excelInput")
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
    statusMessage.innerHTML = message // Changed from textContent to innerHTML
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
        const keyTracker = new Map() // Track first occurrence of each key
        const skippedRows = []
        const duplicateRows = []
        let validPairs = 0

        // Process rows to extract column Q (17) and R (18) data
        // Skip the first row (index 0) assuming it's a header row
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i]
            const rowNumber = row.getAttribute('r') || (i + 1) // Get actual row number from Excel
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

            // Process row if both values exist and are not empty
            if (columnQValue && columnRValue) {
                validPairs++

                // Check for duplicates
                if (mapping.has(columnQValue)) {
                    const firstOccurrence = keyTracker.get(columnQValue)
                    duplicateRows.push({
                        row: rowNumber,
                        columnQ: columnQValue,
                        columnR: columnRValue,
                        previousValue: mapping.get(columnQValue),
                        firstRow: firstOccurrence.row,
                        firstValue: firstOccurrence.value
                    })
                } else {
                    // Track first occurrence
                    keyTracker.set(columnQValue, {
                        row: rowNumber,
                        value: columnRValue
                    })
                }

                mapping.set(columnQValue, columnRValue)
            } else {
                // Track skipped rows
                const missingColumns = []
                if (!columnQValue) missingColumns.push('Q')
                if (!columnRValue) missingColumns.push('R')

                skippedRows.push({
                    row: rowNumber,
                    missing: missingColumns,
                    columnQ: columnQValue || '(empty)',
                    columnR: columnRValue || '(empty)'
                })
            }
        }

        if (mapping.size === 0) {
            showStatus('⚠️ No file mappings found in Excel columns Q and R', 'warning')
        } else {
            // Determine status type based on issues
            let statusType = 'success'
            let statusMsg = `✅ Loaded ${mapping.size} file mappings from Excel`

            const issues = []
            if (skippedRows.length > 0) {
                issues.push(`${skippedRows.length} rows skipped due to missing values`)
            }
            if (duplicateRows.length > 0) {
                issues.push(`${duplicateRows.length} filename conflicts`)
                statusType = 'error' // Show as error if duplicates found
                statusMsg = `⚠️ Loaded ${mapping.size} file mappings from Excel with conflicts`
            }

            if (issues.length > 0) {
                statusMsg += ` (${issues.join(', ')})`

                // Add detailed information
                let details = ''
                if (skippedRows.length > 0) {
                    details += showSkippedRowsDetails(skippedRows, 'Excel')
                }
                if (duplicateRows.length > 0) {
                    details += showDuplicateDetails(duplicateRows, 'Excel')
                }
                statusMsg += details
            }

            showStatus(statusMsg, statusType)
        }

        // Console logging for developers
        if (skippedRows.length > 0) {
            console.log(`📋 Skipped ${skippedRows.length} Excel rows:`)
            skippedRows.forEach(skip => {
                console.log(`Row ${skip.row}: Missing column(s) ${skip.missing.join(', ')} - Q: "${skip.columnQ}", R: "${skip.columnR}"`)
            })
        }

        if (duplicateRows.length > 0) {
            console.log(`📋 Found ${duplicateRows.length} duplicate keys:`)
            duplicateRows.forEach(dup => {
                console.log(`Row ${dup.row}: Key "${dup.columnQ}" already exists (first: row ${dup.firstRow} → "${dup.firstValue}", current: "${dup.columnR}")`)
            })
        }

        return mapping

    } catch (error) {
        console.error('Excel parsing error:', error)
        showStatus(`❌ Failed to parse Excel file: ${error.message}`, 'error')
        return new Map()
    }
}

function showSkippedRowsDetails(skippedRows, fileType) {
    if (skippedRows.length === 0) return ''

    const maxRowsToShow = 5 // Limit display to avoid overwhelming users
    const displayRows = skippedRows.slice(0, maxRowsToShow)
    const hasMore = skippedRows.length > maxRowsToShow

    let details = `<details style="margin-top: 10px; font-size: 0.9rem;">
        <summary style="cursor: pointer; color: #f39c12;">
            ⚠️ ${skippedRows.length} ${fileType} rows skipped - click to view details
        </summary>
        <div style="margin-top: 8px; padding: 8px; background: #fff3cd; border-radius: 4px; font-family: monospace; font-size: 0.85rem;">
            ${displayRows.map(skip => {
        if (skip.reason) {
            return `Line ${skip.row}: ${skip.reason}`
        } else {
            return `Row ${skip.row}: Missing ${skip.missing.join(', ')} - Q: "${skip.columnQ}", R: "${skip.columnR}"`
        }
    }).join('<br>')}
            ${hasMore ? `<br><em>... and ${skippedRows.length - maxRowsToShow} more rows</em>` : ''}
        </div>
    </details>`

    return details
}

function showDuplicateDetails(duplicateRows, fileType) {
    if (duplicateRows.length === 0) return ''

    const maxRowsToShow = 5 // Limit display to avoid overwhelming users
    const displayRows = duplicateRows.slice(0, maxRowsToShow)
    const hasMore = duplicateRows.length > maxRowsToShow

    let details = `<details style="margin-top: 10px; font-size: 0.9rem;">
        <summary style="cursor: pointer; color: #e74c3c; font-weight: 600;">
            ❌ ${duplicateRows.length} conflicting filename(s) - click to view details
        </summary>
        <div style="margin-top: 8px; padding: 8px; background: #f8d7da; border-radius: 4px; font-family: monospace; font-size: 0.85rem; color: #721c24;">
            <strong>The same source file appears multiple times with different target names:</strong><br><br>
            ${displayRows.map(dup => {
        return `<strong>"${dup.columnQ}"</strong> appears in:<br>` +
            `• Row ${dup.firstRow}: rename to "${dup.firstValue}" (first occurrence)<br>` +
            `• Row ${dup.row}: rename to "${dup.columnR}" (duplicate)<br>` +
            `<em>→ Using: "${dup.columnR}" (last occurrence wins)</em><br><br>`
    }).join('')}
            ${hasMore ? `<em>... and ${duplicateRows.length - maxRowsToShow} more conflicts</em><br>` : ''}
            <strong>Action needed:</strong> Remove duplicate rows or decide which target name to use for each source file.
        </div>
    </details>`

    return details
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

            // Store original info for tracking
            fileItem.originalName = originalFileName
            fileItem.finalName = newFileName

            // Create a new file object with the renamed filename if mapping exists
            if (newFileName !== originalFileName) {
                fileItem.file = new File([fileItem.file], newFileName, {
                    type: fileItem.file.type,
                    lastModified: fileItem.file.lastModified
                })
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
            const finalFileName = fileItem.finalName

            try {
                updateProgress(i + 1, heicFiles.length, `Converting ${finalFileName}...`)

                const convertedBlob = await heicTo({
                    blob: fileItem.file,
                    type: 'image/jpeg',
                    quality: 0.8
                })

                // Use the final name (renamed) and convert extension to .jpg
                const baseName = finalFileName.replace(/\.(heic|heif)$/i, '')
                const outputFileName = baseName + '.jpg'
                const zipPath = fileItem.relativePath + outputFileName

                zip.file(zipPath, convertedBlob)
                convertedFiles.push({
                    original: fileItem.originalName,
                    converted: outputFileName,
                    originalSize: fileItem.file.size,
                    convertedSize: convertedBlob.size,
                    wasRenamed: !!fileItem.renamedFrom
                })

                if (fileItem.renamedFrom) {
                    renamedFiles.push({
                        from: fileItem.renamedFrom,
                        to: outputFileName // Show the final output name
                    })
                }

            } catch (error) {
                console.error(`Failed to convert ${finalFileName}:`, error)
                failedFiles.push({
                    name: finalFileName,
                    error: error.message
                })
            }
        }

        // Add non-HEIC image files to ZIP as-is with renamed names
        for (const fileItem of nonHeicFiles) {
            const finalFileName = fileItem.finalName
            const zipPath = fileItem.relativePath + finalFileName
            zip.file(zipPath, fileItem.file)

            if (fileItem.renamedFrom) {
                renamedFiles.push({
                    from: fileItem.renamedFrom,
                    to: finalFileName
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
    // Calculate breakdown
    const totalProcessed = convertedFiles.length + otherFiles + failedFiles.length
    const renamedHeicFiles = convertedFiles.filter(f => f.wasRenamed).length
    const renamedOtherFiles = renamedFiles.length - renamedHeicFiles

    originalImage.innerHTML = `
    <div style="padding: 20px; text-align: center;">
      <div style="font-size: 3rem; margin-bottom: 10px;">📁</div>
      <div style="font-weight: bold; margin-bottom: 10px;">Batch Conversion Results</div>
      <div style="color: #666; font-size: 0.9rem; line-height: 1.5;">
        <div>✅ ${convertedFiles.length} HEIC files converted</div>
        ${otherFiles > 0 ? `<div>📷 ${otherFiles} other image files included</div>` : ''}
        ${renamedFiles.length > 0 ? `<div>🏷️ ${renamedFiles.length} files renamed (${renamedHeicFiles} HEIC, ${renamedOtherFiles} other)</div>` : ''}
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

    if (failedFiles.length > 0) {
        showStatus(
            `⚠️ Processed ${successCount}/${totalProcessed} files. ${failedFiles.length} files failed to convert.`,
            'warning'
        )
    } else {
        let message = `🎉 Successfully processed all ${totalProcessed} files!`
        if (convertedFiles.length > 0) {
            message += ` ${convertedFiles.length} HEIC files converted.`
        }
        if (otherFiles > 0) {
            message += ` ${otherFiles} other files included.`
        }
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

            // Use the renamed filename for the download, strip extension and add .jpg
            const baseName = processedFile.name.replace(/\.[^/.]+$/, '')
            const downloadFileName = baseName + '.jpg'

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