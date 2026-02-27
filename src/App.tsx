import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

type OutputMode = 'objects' | 'arrays'

type ConvertOptions = {
  sheetName: string
  outputMode: OutputMode
  headerRowNumber: number
  skipEmptyRows: boolean
  pretty: boolean
}

function normalizeHeaders(values: unknown[]): string[] {
  const base = values.map((v, i) => {
    const text = String(v ?? '').trim()
    return text.length > 0 ? text : `Column${i + 1}`
  })

  const used = new Map<string, number>()
  return base.map((key) => {
    const count = used.get(key) ?? 0
    used.set(key, count + 1)
    if (count === 0) return key
    return `${key}_${count + 1}`
  })
}

function isRowEmpty(row: unknown[]): boolean {
  return row.every((cell) => {
    if (cell == null) return true
    if (typeof cell === 'string') return cell.trim().length === 0
    return false
  })
}

function clampNumber(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value))
}

function App() {
  const [file, setFile] = useState<File | null>(null)
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [isLoading, setIsLoading] = useState(false)

  const [options, setOptions] = useState<ConvertOptions>({
    sheetName: '',
    outputMode: 'objects',
    headerRowNumber: 1,
    skipEmptyRows: true,
    pretty: true,
  })

  const sheetNames = useMemo(() => workbook?.SheetNames ?? [], [workbook])

  useEffect(() => {
    if (!workbook) return
    setOptions((prev) => ({
      ...prev,
      sheetName: prev.sheetName && workbook.SheetNames.includes(prev.sheetName) ? prev.sheetName : workbook.SheetNames[0] ?? '',
    }))
  }, [workbook])

  async function loadFile(nextFile: File) {
    setIsLoading(true)
    setError(null)
    setFile(nextFile)

    try {
      const buffer = await nextFile.arrayBuffer()
      const wb = XLSX.read(buffer, { type: 'array' })
      if (wb.SheetNames.length === 0) {
        setWorkbook(null)
        setError('Dosyada sayfa bulunamadı.')
        return
      }
      setWorkbook(wb)
    } catch (e) {
      setWorkbook(null)
      const message = e instanceof Error ? e.message : 'Bilinmeyen hata'
      setError(`Dosya okunamadı: ${message}`)
    } finally {
      setIsLoading(false)
    }
  }

  const convertResult = useMemo(() => {
    if (!workbook || !options.sheetName) return null
    const sheet = workbook.Sheets[options.sheetName]
    if (!sheet) return null

    const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: '',
      blankrows: true,
    })

    const headerRowIndex = clampNumber(options.headerRowNumber - 1, 0, Math.max(0, rows.length - 1))

    let data: unknown
    let previewRows: unknown[][]
    let rowCount: number
    if (options.outputMode === 'arrays') {
      const exportRows = options.skipEmptyRows ? rows.filter((r) => !isRowEmpty(r)) : rows
      data = exportRows
      previewRows = exportRows.slice(0, 25)
      rowCount = exportRows.length
    } else {
      const headerRow = rows[headerRowIndex] ?? []
      const headers = normalizeHeaders(headerRow)

      const dataRows = rows.slice(headerRowIndex + 1).filter((r) => (options.skipEmptyRows ? !isRowEmpty(r) : true))
      const objects = dataRows
        .map((r) => {
          const obj: Record<string, unknown> = {}
          for (let i = 0; i < headers.length; i++) {
            obj[headers[i]] = r[i] ?? ''
          }
          return obj
        })

      data = objects
      previewRows = [headerRow, ...dataRows.slice(0, 24)]
      rowCount = objects.length
    }

    const json = JSON.stringify(data, null, options.pretty ? 2 : 0)

    return {
      json,
      previewRows,
      rowCount,
      headerRowIndex,
    }
  }, [options.headerRowNumber, options.outputMode, options.pretty, options.sheetName, options.skipEmptyRows, workbook])

  const canDownload = Boolean(convertResult?.json && file)

  function downloadJson() {
    if (!convertResult?.json) return
    const baseName = file?.name ? file.name.replace(/\.(xlsx|xls|csv)$/i, '') : 'data'
    const outName = `${baseName}.json`
    const blob = new Blob([convertResult.json], { type: 'application/json;charset=utf-8' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = outName
    document.body.appendChild(a)
    a.click()
    a.remove()
    URL.revokeObjectURL(url)
  }

  async function copyJson() {
    if (!convertResult?.json) return
    await navigator.clipboard.writeText(convertResult.json)
  }

  function onPickFile(next: File | null) {
    if (!next) return
    void loadFile(next)
  }

  return (
    <div className="app">
      <header className="header">
        <div className="title">
          <h1>Excel → JSON Dönüştürücü</h1>
          <p>Excel (.xlsx/.xls) veya CSV dosyanı yükle, JSON olarak indir.</p>
        </div>
      </header>

      <main className="layout">
        <section className="panel">
          <div className="block">
            <label className="label" htmlFor="file">
              Dosya
            </label>
            <div
              className="dropzone"
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault()
                const next = e.dataTransfer.files?.[0]
                onPickFile(next ?? null)
              }}
            >
              <input
                id="file"
                type="file"
                accept=".xlsx,.xls,.csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv"
                onChange={(e) => onPickFile(e.target.files?.[0] ?? null)}
              />
              <div className="dropzoneText">
                <div className="fileName">{file?.name ?? 'Dosya seç veya sürükleyip bırak'}</div>
                <div className="hint">Maksimum hız için dönüşüm tarayıcında yapılır.</div>
              </div>
            </div>
            {isLoading ? <div className="status">Okunuyor…</div> : null}
            {error ? <div className="error">{error}</div> : null}
          </div>

          <div className="block">
            <div className="grid2">
              <div>
                <label className="label" htmlFor="sheet">
                  Sayfa
                </label>
                <select
                  id="sheet"
                  value={options.sheetName}
                  disabled={!workbook || sheetNames.length === 0}
                  onChange={(e) => setOptions((p) => ({ ...p, sheetName: e.target.value }))}
                >
                  {sheetNames.map((name) => (
                    <option key={name} value={name}>
                      {name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="label" htmlFor="mode">
                  Çıktı
                </label>
                <select
                  id="mode"
                  value={options.outputMode}
                  disabled={!workbook}
                  onChange={(e) => setOptions((p) => ({ ...p, outputMode: e.target.value as OutputMode }))}
                >
                  <option value="objects">Nesne dizisi (başlık satırı)</option>
                  <option value="arrays">Satır dizisi (array of arrays)</option>
                </select>
              </div>
            </div>

            <div className="grid2">
              <div>
                <label className="label" htmlFor="headerRow">
                  Başlık satırı (1’den başlar)
                </label>
                <input
                  id="headerRow"
                  type="number"
                  min={1}
                  value={options.headerRowNumber}
                  disabled={!workbook || options.outputMode !== 'objects'}
                  onChange={(e) => setOptions((p) => ({ ...p, headerRowNumber: Number(e.target.value || 1) }))}
                />
              </div>

              <div className="toggles">
                <label className="toggle">
                  <input
                    type="checkbox"
                    checked={options.skipEmptyRows}
                    disabled={!workbook}
                    onChange={(e) => setOptions((p) => ({ ...p, skipEmptyRows: e.target.checked }))}
                  />
                  Boş satırları atla
                </label>
                <label className="toggle">
                  <input
                    type="checkbox"
                    checked={options.pretty}
                    disabled={!workbook}
                    onChange={(e) => setOptions((p) => ({ ...p, pretty: e.target.checked }))}
                  />
                  Girintili JSON
                </label>
              </div>
            </div>

            <div className="actions">
              <button type="button" disabled={!canDownload} onClick={downloadJson}>
                JSON indir
              </button>
              <button type="button" className="secondary" disabled={!convertResult?.json} onClick={() => void copyJson()}>
                JSON kopyala
              </button>
            </div>

            {convertResult ? (
              <div className="meta">
                <div>Toplam satır: {convertResult.rowCount}</div>
                {options.outputMode === 'objects' ? <div>Başlık satırı: {convertResult.headerRowIndex + 1}</div> : null}
              </div>
            ) : null}
          </div>
        </section>

        <section className="panel">
          <div className="block">
            <div className="labelRow">
              <div className="label">Önizleme</div>
              <div className="hint">İlk 25 satır</div>
            </div>
            {convertResult?.previewRows?.length ? (
              <div className="tableWrap" role="region" aria-label="Önizleme tablosu" tabIndex={0}>
                <table>
                  <tbody>
                    {convertResult.previewRows.map((row, rIndex) => (
                      <tr key={rIndex}>
                        {row.map((cell, cIndex) => (
                          <td key={cIndex}>{String(cell ?? '')}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="empty">Dosya yükleyince burada önizleme göreceksin.</div>
            )}
          </div>

          <div className="block">
            <div className="labelRow">
              <div className="label">JSON</div>
              <div className="hint">İndirmeden önce kontrol edebilirsin</div>
            </div>
            <textarea readOnly value={convertResult?.json ?? ''} placeholder="JSON çıktısı burada görünecek" />
          </div>
        </section>
      </main>
    </div>
  )
}

export default App
