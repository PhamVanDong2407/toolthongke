import { useState, useRef } from 'react'
import * as XLSX from 'xlsx'

function parseDate(val) {
  if (!val) return null
  if (val instanceof Date) return val
  if (typeof val === 'number') {
    // Excel serial date
    const utcDays = Math.floor(val - 25569)
    const date = new Date(utcDays * 86400 * 1000)
    return date
  }
  if (typeof val === 'string') {
    const d = new Date(val)
    if (!isNaN(d)) return d
  }
  return null
}

export default function App() {
  const [fileName, setFileName] = useState('')
  const [error, setError] = useState('')
  const [workbookData, setWorkbookData] = useState(null)
  const [months, setMonths] = useState([])
  const [years, setYears] = useState([])
  const [selectedMonth, setSelectedMonth] = useState('')
  const [selectedYear, setSelectedYear] = useState('')
  const [result, setResult] = useState(null)
  const [isDragging, setIsDragging] = useState(false)
  const fileInputRef = useRef(null)

  const handleFile = (file) => {
    setError('')
    setResult(null)
    setFileName('')
    setWorkbookData(null)

    if (!file.name.match(/\.xlsx?$/i)) {
      setError('Vui lòng chọn file Excel (.xlsx)')
      return
    }

    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result)
        const wb = XLSX.read(data, { type: 'array', cellDates: true })

        const sheetNames = wb.SheetNames
        const hasFollowers = sheetNames.includes('Người theo dõi')
        const hasViews = sheetNames.includes('Lượt tương tác')

        if (!hasFollowers && !hasViews) {
          setError('File không đúng định dạng Telegram Stats. Cần có sheet "Người theo dõi" hoặc "Lượt tương tác".')
          return
        }

        const followers = hasFollowers
          ? XLSX.utils.sheet_to_json(wb.Sheets['Người theo dõi'], { defval: 0 })
          : []
        const views = hasViews
          ? XLSX.utils.sheet_to_json(wb.Sheets['Lượt tương tác'], { defval: 0 })
          : []

        const availableMonths = new Set()
        const availableYears = new Set()

        ;[...followers, ...views].forEach((row) => {
          const d = parseDate(row.Date || row.date)
          if (d) {
            availableMonths.add(d.getMonth() + 1)
            availableYears.add(d.getFullYear())
          }
        })

        const sortedMonths = [...availableMonths].sort((a, b) => a - b)
        const sortedYears = [...availableYears].sort((a, b) => b - a)

        setWorkbookData({ followers, views })
        setMonths(sortedMonths)
        setYears(sortedYears)
        setSelectedMonth(sortedMonths[sortedMonths.length - 1] || '')
        setSelectedYear(sortedYears[0] || '')
        setFileName(file.name)
      } catch (err) {
        console.error(err)
        setError('Không đọc được file. Vui lòng kiểm tra lại file xuất từ Telegram.')
      }
    }
    reader.readAsArrayBuffer(file)
  }

  const onDrop = (e) => {
    e.preventDefault()
    setIsDragging(false)
    if (e.dataTransfer.files.length) {
      handleFile(e.dataTransfer.files[0])
    }
  }

  const calculate = () => {
    if (!workbookData) return

    const month = parseInt(selectedMonth)
    const year = parseInt(selectedYear)

    let totalViews = 0
    let totalJoined = 0
    let totalLeft = 0
    let daysWithData = 0
    let daysWithViews = 0

    workbookData.followers.forEach((row) => {
      const d = parseDate(row.Date || row.date)
      if (d && d.getMonth() + 1 === month && d.getFullYear() === year) {
        totalJoined += Number(row.Joined || row.joined || 0)
        totalLeft += Number(row.Left || row.left || 0)
        daysWithData++
      }
    })

    workbookData.views.forEach((row) => {
      const d = parseDate(row.Date || row.date)
      if (d && d.getMonth() + 1 === month && d.getFullYear() === year) {
        totalViews += Number(row.Views || row.views || 0)
        daysWithViews++
      }
    })

    const net = totalJoined - totalLeft
    const avgViews = daysWithViews ? Math.round(totalViews / daysWithViews) : 0

    setResult({
      month,
      year,
      totalViews,
      totalJoined,
      totalLeft,
      net,
      daysWithData,
      daysWithViews,
      avgViews,
    })
  }

  return (
    <div className="container">
      <header>
        <h1>📊 Telegram Channel Stats</h1>
        <p>Upload file xuất thống kê (.xlsx) → Tính tổng View, Follow mới & Follow rời theo tháng</p>
      </header>

      {error && <div className="error">{error}</div>}

      <div
        className={`upload-zone ${isDragging ? 'dragover' : ''} ${fileName ? 'has-file' : ''}`}
        onClick={() => fileInputRef.current?.click()}
        onDragOver={(e) => {
          e.preventDefault()
          setIsDragging(true)
        }}
        onDragLeave={() => setIsDragging(false)}
        onDrop={onDrop}
      >
        <div className="upload-icon">📁</div>
        <h3>Kéo thả file .xlsx vào đây</h3>
        <p>hoặc bấm để chọn file từ máy</p>
        {fileName && <div className="file-name">✅ {fileName}</div>}
        <input
          type="file"
          ref={fileInputRef}
          accept=".xlsx,.xls"
          style={{ display: 'none' }}
          onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])}
        />
      </div>

      {workbookData && (
        <div className="controls">
          <div className="form-group">
            <label>Chọn tháng</label>
            <select value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)}>
              {months.map((m) => (
                <option key={m} value={m}>
                  Tháng {m}
                </option>
              ))}
            </select>
          </div>
          <div className="form-group">
            <label>Chọn năm</label>
            <select value={selectedYear} onChange={(e) => setSelectedYear(e.target.value)}>
              {years.map((y) => (
                <option key={y} value={y}>
                  {y}
                </option>
              ))}
            </select>
          </div>
          <button onClick={calculate}>Tính toán</button>
        </div>
      )}

      {result && (
        <div>
          <div className="month-title">
            Tháng {result.month}/{result.year}
          </div>
          <div className="cards">
            <div className="card views">
              <div className="label">Tổng lượt xem</div>
              <div className="value">{result.totalViews.toLocaleString('vi-VN')}</div>
            </div>
            <div className="card joined">
              <div className="label">Follow mới</div>
              <div className="value">{result.totalJoined.toLocaleString('vi-VN')}</div>
            </div>
            <div className="card left">
              <div className="label">Follow rời đi</div>
              <div className="value">{result.totalLeft.toLocaleString('vi-VN')}</div>
            </div>
            <div className="card net">
              <div className="label">Tăng trưởng ròng</div>
              <div className="value">
                {(result.net >= 0 ? '+' : '') + result.net.toLocaleString('vi-VN')}
              </div>
            </div>
          </div>
          <div className="extra">
            <strong>Chi tiết:</strong>
            <br />
            • Số ngày có dữ liệu follow: <strong>{result.daysWithData}</strong> ngày
            <br />
            • Số ngày có dữ liệu view: <strong>{result.daysWithViews}</strong> ngày
            <br />
            • Trung bình view/ngày: <strong>{result.avgViews.toLocaleString('vi-VN')}</strong>
          </div>
        </div>
      )}

      <footer>
        Dữ liệu được xử lý hoàn toàn trên trình duyệt của bạn · Không upload lên server
      </footer>
    </div>
  )
}