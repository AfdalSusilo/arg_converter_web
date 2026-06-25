'use client'

import { useState, useCallback, useRef } from 'react'
import { Upload, FileSpreadsheet, FileText, X, CheckCircle, AlertCircle, Info, Table } from 'lucide-react'
import Papa from 'papaparse'
import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'

type LogType = 'auto' | 'behavior' | 'gui' | 'npc_chat' | 'unknown'

interface ParsedRow {
  [key: string]: string | number | null
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null)
  const [data, setData] = useState<ParsedRow[]>([])
  const [columns, setColumns] = useState<string[]>([])
  const [logType, setLogType] = useState<LogType>('auto')
  const [status, setStatus] = useState<{ type: 'success' | 'error' | 'info'; message: string } | null>(null)
  const [isProcessing, setIsProcessing] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)

  // Flatten JSON string columns
  const flattenData = useCallback((df: ParsedRow[]): ParsedRow[] => {
    return df.map(row => {
      const newRow: ParsedRow = {}
      for (const [key, value] of Object.entries(row)) {
        if (typeof value === 'string' && (value.startsWith('[') || value.startsWith('{'))) {
          try {
            const parsed = JSON.parse(value)
            if (Array.isArray(parsed)) {
              parsed.forEach((item, idx) => {
                if (typeof item === 'object') {
                  Object.entries(item).forEach(([k, v]) => {
                    newRow[`${key}_${k}_${idx}`] = v as string | number | null
                  })
                } else {
                  newRow[`${key}_${idx}`] = item as string | number | null
                }
              })
            } else if (typeof parsed === 'object' && parsed !== null) {
              Object.entries(parsed).forEach(([k, v]) => {
                newRow[`${key}_${k}`] = v as string | number | null
              })
            } else {
              newRow[key] = value
            }
          } catch {
            newRow[key] = value
          }
        } else {
          newRow[key] = value
        }
      }
      return newRow
    })
  }, [])

  // Auto-detect log type based on columns
  const detectLogType = useCallback((cols: string[]): LogType => {
    const colsLower = cols.map(c => c.toLowerCase())
    
    if (colsLower.some(c => c.includes('behavior') || c.includes('mouse_events') || c.includes('keystrokes'))) {
      return 'behavior'
    }
    if (colsLower.some(c => c.includes('gui') || c.includes('input_data') || c.includes('widget'))) {
      return 'gui'
    }
    if (colsLower.some(c => c.includes('chat') || c.includes('message') || c.includes('npc_chat'))) {
      return 'npc_chat'
    }
    return 'unknown'
  }, [])

  // Handle file selection
  const handleFileChange = useCallback((selectedFile: File) => {
    if (!selectedFile || !selectedFile.name.endsWith('.csv')) {
      setStatus({ type: 'error', message: 'Please select a valid CSV file' })
      return
    }

    setFile(selectedFile)
    setStatus({ type: 'info', message: 'Processing CSV file...' })
    setIsProcessing(true)

    Papa.parse(selectedFile, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        const parsedData = results.data as ParsedRow[]
        const cols = results.meta.fields || []
        
        setData(parsedData)
        setColumns(cols)
        
        const detected = detectLogType(cols)
        setLogType(detected)
        
        setStatus({
          type: 'success',
          message: `Loaded ${parsedData.length} rows. ${cols.length} columns detected.`
        })
        setIsProcessing(false)
      },
      error: (error) => {
        setStatus({ type: 'error', message: `Error parsing CSV: ${error.message}` })
        setIsProcessing(false)
      }
    })
  }, [detectLogType])

  // Drag and drop handlers
  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
  }, [])

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
    
    const droppedFile = e.dataTransfer.files[0]
    if (droppedFile) {
      handleFileChange(droppedFile)
    }
  }, [handleFileChange])

  // Remove file
  const handleRemoveFile = useCallback(() => {
    setFile(null)
    setData([])
    setColumns([])
    setStatus(null)
    if (fileInputRef.current) {
      fileInputRef.current.value = ''
    }
  }, [])

  // Export to Excel
  const exportToExcel = useCallback(() => {
    if (data.length === 0) return

    try {
      setStatus({ type: 'info', message: 'Generating Excel file...' })
      
      const flattenedData = flattenData(data)
      const ws = XLSX.utils.json_to_sheet(flattenedData)
      const wb = XLSX.utils.book_new()
      
      // Auto-fit columns
      const colWidths = Object.keys(flattenedData[0] || {}).map(key => {
        let maxLen = key.length
        flattenedData.forEach(row => {
          const val = String(row[key] || '')
          if (val.length > maxLen) maxLen = val.length
        })
        return { wch: Math.min(maxLen + 2, 50) }
      })
      ws['!cols'] = colWidths
      
      XLSX.utils.book_append_sheet(wb, ws, 'ARG_Report')
      
      const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
      
      const fileName = file?.name.replace('.csv', '_report.xlsx') || 'arg_report.xlsx'
      saveAs(blob, fileName)
      
      setStatus({ type: 'success', message: 'Excel file downloaded successfully!' })
    } catch (error) {
      setStatus({ type: 'error', message: `Error exporting to Excel: ${error}` })
    }
  }, [data, file, flattenData])

  // Export to Word (as formatted HTML that can be opened in Word)
  const exportToWord = useCallback(() => {
    if (data.length === 0) return

    try {
      setStatus({ type: 'info', message: 'Generating Word document...' })
      
      const flattenedData = flattenData(data)
      
      // Detect player_name column
      let playerCol = ''
      for (const col of ['player_name', 'PlayerName', 'player', 'Player', 'username', 'Username']) {
        if (columns.includes(col)) {
          playerCol = col
          break
        }
      }

      // Build HTML content for Word
      let html = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office"
              xmlns:w="urn:schemas-microsoft-com:office:word"
              xmlns="http://www.w3.org/TR/REC-html40">
        <head>
          <meta charset="utf-8">
          <title>ARG Log Report</title>
          <style>
            body { font-family: Calibri, sans-serif; margin: 40px; }
            h1 { color: #1E3A5F; border-bottom: 2px solid #87CEEB; padding-bottom: 10px; }
            h2 { color: #1E3A5F; margin-top: 30px; }
            .meta { color: #666; margin-bottom: 30px; }
            .record { margin-bottom: 20px; padding: 15px; border: 1px solid #E6F3FF; border-radius: 8px; background: #F9FAFB; }
            .label { font-weight: bold; color: #1E3A5F; }
            .value { margin-left: 10px; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th { background: #1E3A5F; color: white; padding: 10px; text-align: left; }
            td { padding: 10px; border: 1px solid #E5E7EB; }
            tr:nth-child(even) { background: #F9FAFB; }
          </style>
        </head>
        <body>
          <h1>ARG Log Report</h1>
          <div class="meta">
            <p><strong>Generated:</strong> ${new Date().toLocaleString()}</p>
            <p><strong>Log Type:</strong> ${logType === 'auto' ? 'Auto-Detected' : logType}</p>
            <p><strong>Total Records:</strong> ${data.length}</p>
          </div>
      `

      if (playerCol) {
        // Group by player
        const grouped: Record<string, ParsedRow[]> = {}
        flattenedData.forEach(row => {
          const player = String(row[playerCol] || 'Unknown')
          if (!grouped[player]) grouped[player] = []
          grouped[player].push(row)
        })

        Object.entries(grouped).forEach(([player, rows]) => {
          html += `<h2>${player}</h2>`
          rows.forEach(row => {
            html += '<div class="record">'
            Object.entries(row).forEach(([key, value]) => {
              if (key !== playerCol && value !== null && value !== '') {
                html += `<p><span class="label">${key}:</span><span class="value">${value}</span></p>`
              }
            })
            html += '</div>'
          })
        })
      } else {
        // Simple table view
        html += '<table><thead><tr>'
        Object.keys(flattenedData[0] || {}).forEach(col => {
          html += `<th>${col}</th>`
        })
        html += '</tr></thead><tbody>'
        flattenedData.slice(0, 100).forEach(row => {
          html += '<tr>'
          Object.values(row).forEach(val => {
            html += `<td>${val ?? ''}</td>`
          })
          html += '</tr>'
        })
        html += '</tbody></table>'
      }

      html += '</body></html>'

      const blob = new Blob([html], { type: 'application/msword' })
      const fileName = file?.name.replace('.csv', '_report.doc') || 'arg_report.doc'
      saveAs(blob, fileName)
      
      setStatus({ type: 'success', message: 'Word document downloaded successfully!' })
    } catch (error) {
      setStatus({ type: 'error', message: `Error exporting to Word: ${error}` })
    }
  }, [data, file, logType, columns, flattenData])

  return (
    <main className="container">
      <header className="header">
        <h1>ARG Converter</h1>
        <p>Convert CSV Logs to Excel & Word Reports</p>
      </header>

      {/* File Upload Card */}
      <section className="card">
        <h2 className="card-title">Select CSV File</h2>
        
        <div
          className="file-upload"
          onDragOver={handleDragOver}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
        >
          <input
            ref={fileInputRef}
            type="file"
            accept=".csv"
            onChange={(e) => e.target.files?.[0] && handleFileChange(e.target.files[0])}
          />
          <div className="file-upload-icon">
            <Upload size={48} />
          </div>
          <p className="file-upload-text">
            Drag and drop your CSV file here, or click to browse
          </p>
          <p className="file-upload-hint">
            Supports .csv files with messy JSON log data
          </p>
        </div>

        {file && (
          <div className="file-info">
            <FileSpreadsheet className="file-info-icon" size={24} />
            <div className="file-info-details">
              <p className="file-info-name">{file.name}</p>
              <p className="file-info-size">
                {(file.size / 1024).toFixed(2)} KB
              </p>
            </div>
            <button className="file-info-remove" onClick={handleRemoveFile}>
              <X size={20} />
            </button>
          </div>
        )}
      </section>

      {/* Log Type Selection */}
      <section className="card">
        <h2 className="card-title">Log Type</h2>
        <div className="log-type-selector">
          {[
            { value: 'auto', label: 'Auto-Detect' },
            { value: 'behavior', label: 'Behavior Logs' },
            { value: 'gui', label: 'GUI Logs' },
            { value: 'npc_chat', label: 'NPC Chat' },
          ].map((option) => (
            <div key={option.value} className="log-type-option">
              <input
                type="radio"
                id={option.value}
                name="logType"
                value={option.value}
                checked={logType === option.value}
                onChange={(e) => setLogType(e.target.value as LogType)}
              />
              <label htmlFor={option.value} className="log-type-label">
                {option.label}
              </label>
            </div>
          ))}
        </div>
      </section>

      {/* Export Buttons */}
      <section className="card">
        <h2 className="card-title">Export Report</h2>
        <div className="export-buttons">
          <button
            className="btn btn-primary"
            onClick={exportToExcel}
            disabled={data.length === 0 || isProcessing}
          >
            <FileSpreadsheet size={20} />
            Export to Excel
          </button>
          <button
            className="btn btn-secondary"
            onClick={exportToWord}
            disabled={data.length === 0 || isProcessing}
          >
            <FileText size={20} />
            Export to Word
          </button>
        </div>
      </section>

      {/* Status Message */}
      {status && (
        <div className={`status-message status-${status.type}`}>
          {status.type === 'success' && <CheckCircle size={20} style={{ marginRight: 8, verticalAlign: 'middle' }} />}
          {status.type === 'error' && <AlertCircle size={20} style={{ marginRight: 8, verticalAlign: 'middle' }} />}
          {status.type === 'info' && <Info size={20} style={{ marginRight: 8, verticalAlign: 'middle' }} />}
          {status.message}
        </div>
      )}

      {/* Data Preview */}
      {data.length > 0 && (
        <section className="card preview-section">
          <h2 className="preview-title">
            <Table size={20} style={{ marginRight: 8, verticalAlign: 'middle' }} />
            Data Preview (First 10 rows)
          </h2>
          <div className="preview-table">
            <table>
              <thead>
                <tr>
                  {columns.slice(0, 10).map((col) => (
                    <th key={col}>{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {data.slice(0, 10).map((row, idx) => (
                  <tr key={idx}>
                    {columns.slice(0, 10).map((col) => (
                      <td key={col} title={String(row[col] || '')}>
                        {String(row[col] || '')}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      <footer className="footer">
        <p>
          ARG Converter Web - Convert CSV logs to Excel & Word reports
        </p>
        <p style={{ marginTop: 8 }}>
          <a href="https://github.com/AfdalSusilo/arg_converter_web" target="_blank" rel="noopener noreferrer">
            View on GitHub
          </a>
        </p>
      </footer>
    </main>
  )
}
