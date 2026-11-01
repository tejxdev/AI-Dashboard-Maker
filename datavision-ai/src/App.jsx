import React, { useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import {
  Activity, ArrowDownToLine, BarChart3, Bell, CheckCircle2, ChevronDown,
  ChevronRight, CircleHelp, Database, FileSpreadsheet, FileUp, Filter,
  FolderOpen, LayoutDashboard, Lightbulb, MessageCircle, Moon, MoreHorizontal, Search,
  Printer, Send, Settings2, Sparkles, Sun, Table2, TrendingUp, UploadCloud, UserRound, X, Zap
} from 'lucide-react'
import {
  ResponsiveContainer, AreaChart, Area, BarChart, Bar, CartesianGrid,
  XAxis, YAxis, Tooltip, PieChart, Pie, Cell, Legend, LineChart, Line,
  ScatterChart, Scatter, ZAxis, Treemap, FunnelChart, Funnel, LabelList
} from 'recharts'

const sampleRows = [
  { Date:'2026-01-05', Region:'North', Category:'Electronics', Product:'Laptop', Sales:125000, Profit:25000, Orders:18 },
  { Date:'2026-01-12', Region:'South', Category:'Furniture', Product:'Office Chair', Sales:48000, Profit:9200, Orders:12 },
  { Date:'2026-02-02', Region:'East', Category:'Electronics', Product:'Headphones', Sales:36000, Profit:8500, Orders:25 },
  { Date:'2026-02-18', Region:'West', Category:'Clothing', Product:'Jacket', Sales:27000, Profit:6300, Orders:21 },
  { Date:'2026-03-03', Region:'North', Category:'Furniture', Product:'Desk', Sales:76000, Profit:14800, Orders:14 },
  { Date:'2026-03-20', Region:'South', Category:'Electronics', Product:'Smartphone', Sales:148000, Profit:29600, Orders:31 },
  { Date:'2026-04-08', Region:'East', Category:'Clothing', Product:'Shoes', Sales:42000, Profit:10500, Orders:28 },
  { Date:'2026-04-22', Region:'West', Category:'Furniture', Product:'Bookshelf', Sales:59000, Profit:11800, Orders:10 },
  { Date:'2026-05-11', Region:'North', Category:'Electronics', Product:'Monitor', Sales:91000, Profit:18200, Orders:16 },
  { Date:'2026-05-29', Region:'South', Category:'Clothing', Product:'T-Shirt', Sales:22000, Profit:5500, Orders:34 },
  { Date:'2026-06-07', Region:'East', Category:'Electronics', Product:'Tablet', Sales:83000, Profit:16600, Orders:19 },
  { Date:'2026-06-19', Region:'West', Category:'Furniture', Product:'Sofa', Sales:112000, Profit:22400, Orders:8 }
]
const COLORS = ['#8b7cff','#37c9a5','#55a7ff','#f5b85c','#ef7d9b','#9b8cff']
const fmt = n => new Intl.NumberFormat('en-IN', { maximumFractionDigits: 0 }).format(Number(n || 0))
const isNumeric = v => v !== null && v !== '' && Number.isFinite(Number(String(v).replace(/,/g,'')))
const cleanNum = v => Number(String(v ?? 0).replace(/,/g,'')) || 0
const isMissingValue = value => value == null || (typeof value === 'string' && value.trim() === '')
const titleCase = s => String(s).replace(/[_-]/g,' ').replace(/\b\w/g, c => c.toUpperCase())
const findColumn = (headers, terms) => headers.find(h => terms.some(t => h.toLowerCase().replace(/[_\s-]/g,'').includes(t)))
const normalizeText = value => String(value).toLowerCase().replace(/[^a-z0-9]+/g,' ').trim()
const normalizeDateValue = value => {
  if (isMissingValue(value) || typeof value === 'number') return false
  const text = String(value).trim()
  const isoMatch = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:T.*)?$/)
  if (isoMatch) {
    const [,year,month,day] = isoMatch.map(Number)
    const date = new Date(Date.UTC(year,month-1,day))
    if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month-1 || date.getUTCDate() !== day) return null
    return `${String(year).padStart(4,'0')}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`
  }
  const datePattern = /^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$/.test(text)
  if (!datePattern) return null
  const date = new Date(text)
  if (Number.isNaN(date.getTime())) return null
  return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`
}
const isDateValue = value => Boolean(normalizeDateValue(value))

function getDuplicateRowCount(rows, headers) {
  const seen = new Set()
  let duplicates = 0
  rows.forEach(row => {
    const key = JSON.stringify(headers.map(header => isMissingValue(row[header]) ? '' : String(row[header]).trim()))
    if (seen.has(key)) duplicates += 1
    else seen.add(key)
  })
  return duplicates
}

function readUploadFile(file) {
  const extension = file.name.split('.').pop()?.toLowerCase()
  if (!['csv','xlsx','xls'].includes(extension)) {
    return Promise.reject(new Error(`${file.name}: unsupported file type. Choose CSV, XLSX or XLS.`))
  }

  return new Promise((resolve,reject) => {
    const reader = new FileReader()
    reader.onerror = () => reject(new Error(`${file.name}: could not read this file.`))
    reader.onload = event => {
      try {
        const workbook = XLSX.read(event.target.result, {
          type:extension === 'csv' ? 'string' : 'array',
          cellDates:true
        })
        const firstSheetName = workbook.SheetNames[0]
        if (!firstSheetName) throw new Error('No worksheet found.')
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], {defval:'',raw:false})
          .filter(row => Object.values(row).some(value => !isMissingValue(value)))
        if (!rows.length) throw new Error('No data rows found in the first worksheet.')
        resolve({fileName:file.name,sheetName:firstSheetName,rows})
      } catch (error) {
        reject(new Error(`${file.name}: ${error.message}`))
      }
    }
    if (extension === 'csv') reader.readAsText(file)
    else reader.readAsArrayBuffer(file)
  })
}

function combineDatasets(datasets) {
  const columns = []
  const columnByKey = new Map()
  datasets.forEach(({rows}) => rows.forEach(row => Object.keys(row).forEach(header => {
    const cleanHeader = String(header).replace(/^\uFEFF/,'').trim()
    if (!cleanHeader) return
    const key = cleanHeader.toLowerCase()
    if (!columnByKey.has(key)) {
      columnByKey.set(key,cleanHeader)
      columns.push(cleanHeader)
    }
  })))

  const mergedRows = datasets.flatMap(({rows}) => rows.map(row => {
    const alignedRow = Object.fromEntries(columns.map(header => [header,'']))
    Object.entries(row).forEach(([header,value]) => {
      const cleanHeader = String(header).replace(/^\uFEFF/,'').trim()
      const canonicalHeader = columnByKey.get(cleanHeader.toLowerCase())
      if (canonicalHeader) alignedRow[canonicalHeader] = value
    })
    return alignedRow
  }))

  return {headers:columns,rows:mergedRows}
}

function answerDataQuestion(question, rows, headers) {
  if (!rows.length) return 'There are no matching records to analyze. Try changing the dashboard filters or asking about a different value.'

  const normalizedQuestion = ` ${normalizeText(question)} `
  const numericCols = headers.filter(h => rows.some(r => isNumeric(r[h])))
  const categoricalCols = headers.filter(h => !numericCols.includes(h))
  const hasTerm = terms => terms.some(term => normalizedQuestion.includes(` ${normalizeText(term)} `))
  const matchingColumn = columns => [...columns]
    .filter(column => normalizedQuestion.includes(` ${normalizeText(column)} `))
    .sort((a,b) => normalizeText(b).length - normalizeText(a).length)[0]

  const metricCol = matchingColumn(numericCols) ||
    (hasTerm(['profit','earnings','margin']) && findColumn(numericCols,['profit','earning','margin'])) ||
    (hasTerm(['sales','revenue','turnover','amount']) && findColumn(numericCols,['sales','revenue','amount','total'])) ||
    (hasTerm(['orders','order','quantity','qty','units']) && findColumn(numericCols,['orders','quantity','qty','count'])) ||
    (hasTerm(['total','highest','lowest','top','bottom','most','least']) && findColumn(numericCols,['sales','revenue','amount','total']))

  const groupAliases = [
    { terms:['product','item'], columns:['product','item'] },
    { terms:['region','state','city','country','location'], columns:['region','state','city','country','location'] },
    { terms:['category','segment','department','type'], columns:['category','segment','department','type'] },
    { terms:['customer','client'], columns:['customer','client'] }
  ]
  const groupCol = matchingColumn(categoricalCols) ||
    groupAliases.find(group => hasTerm(group.terms))?.columns
      .map(term => findColumn(categoricalCols,[term]))
      .find(Boolean) ||
    (hasTerm(['which','highest','lowest','top','bottom','most','least']) &&
      findColumn(categoricalCols,['product','item','region','state','city','country','category','segment','customer']))

  const valueFilters = categoricalCols.map(column => [
    column,
    [...new Set(rows.map(row => row[column]).filter(value => value !== '' && value != null))]
      .filter(value => normalizedQuestion.includes(` ${normalizeText(value)} `))
  ]).filter(([,values]) => values.length)
  const scopedRows = valueFilters.length
    ? rows.filter(row => valueFilters.every(([column,values]) => values.some(value => String(row[column]) === String(value))))
    : rows
  if (!scopedRows.length) return 'I couldn’t find records matching that value in the current data. Check the spelling or try a value shown in Data Explorer.'

  const recordCountRequested = hasTerm(['how many records','number of records','rows','record count']) ||
    (hasTerm(['how many','count']) && !hasTerm(['orders','order','quantity','qty','units']))
  if (recordCountRequested) return `There are ${fmt(scopedRows.length)} matching records in the current data.`

  const averageRequested = hasTerm(['average','avg','mean'])
  const lowestRequested = hasTerm(['lowest','least','smallest','minimum','min','bottom','worst'])
  const highestRequested = hasTerm(['highest','most','largest','maximum','max','top','best'])
  const groupedRequest = Boolean(groupCol) && (highestRequested || lowestRequested || hasTerm(['which','by','per']))

  if (!metricCol) {
    return `I couldn’t identify a numeric column for that question. Available numeric columns: ${numericCols.map(titleCase).join(', ') || 'none'}. Try asking for a total, average, highest, or lowest value.`
  }

  const valueFor = values => averageRequested
    ? values.reduce((sum,value) => sum + cleanNum(value),0) / values.length
    : values.reduce((sum,value) => sum + cleanNum(value),0)

  if (groupedRequest) {
    const groups = new Map()
    scopedRows.forEach(row => {
      const name = String(row[groupCol] ?? '').trim() || 'Unspecified'
      if (!groups.has(name)) groups.set(name, [])
      groups.get(name).push(row[metricCol])
    })
    const ranked = [...groups].map(([name,values]) => ({ name, value:valueFor(values) }))
      .sort((a,b) => lowestRequested ? a.value-b.value : b.value-a.value)
    if (!ranked.length) return `There is no ${titleCase(metricCol)} data to compare in these records.`
    const result = ranked[0]
    const rank = lowestRequested ? 'lowest' : 'highest'
    const aggregation = averageRequested ? 'average' : 'total'
    return `${result.name} has the ${rank} ${aggregation} ${titleCase(metricCol)} at ${fmt(result.value)} (${titleCase(groupCol)}; ${fmt(scopedRows.length)} matching records).`
  }

  if (highestRequested || lowestRequested) {
    const values = scopedRows.map(row => cleanNum(row[metricCol]))
    const result = lowestRequested ? Math.min(...values) : Math.max(...values)
    return `The ${lowestRequested ? 'lowest' : 'highest'} ${titleCase(metricCol)} in ${fmt(scopedRows.length)} matching records is ${fmt(result)}.`
  }
  if (averageRequested) {
    return `The average ${titleCase(metricCol)} across ${fmt(scopedRows.length)} matching records is ${fmt(valueFor(scopedRows.map(row => row[metricCol])))}.`
  }

  return `The total ${titleCase(metricCol)} across ${fmt(scopedRows.length)} matching records is ${fmt(scopedRows.reduce((sum,row) => sum + cleanNum(row[metricCol]),0))}.`
}

function App() {
  const [rows, setRows] = useState(sampleRows)
  const [dataHeaders, setDataHeaders] = useState(Object.keys(sampleRows[0]))
  const [fileName, setFileName] = useState('Sample_Sales_Data.xlsx')
  const [loadedFileNames, setLoadedFileNames] = useState(['Sample_Sales_Data.xlsx'])
  const [isLoadingFiles, setIsLoadingFiles] = useState(false)
  const [active, setActive] = useState('Overview')
  const [dark, setDark] = useState(true)
  const [query, setQuery] = useState('')
  const [region, setRegion] = useState('All regions')
  const [category, setCategory] = useState('All categories')
  const [notice, setNotice] = useState('Sample data loaded — upload your own Excel or CSV file to get started.')
  const [sheetNames, setSheetNames] = useState([])
  const [sheet, setSheet] = useState('')
  const fileRef = useRef(null)

  const headers = dataHeaders
  const numericCols = useMemo(() => headers.filter(h => rows.some(r => isNumeric(r[h]))), [headers, rows])
  const categoricalCols = useMemo(() => headers.filter(h => !numericCols.includes(h)), [headers, numericCols])
  const salesCol = findColumn(numericCols, ['sales','revenue','amount','total'])
  const profitCol = findColumn(numericCols, ['profit','earning','margin'])
  const orderCol = findColumn(numericCols, ['orders','quantity','qty','count'])
  const dateCol = findColumn(headers, ['date','month','year'])
  const categoryCol = findColumn(categoricalCols, ['category','type','segment','department']) || categoricalCols[0]
  const regionCol = findColumn(categoricalCols, ['region','state','city','country','location'])
  const productCol = findColumn(categoricalCols, ['product','item','customer','name'])

  const regions = regionCol ? ['All regions', ...new Set(rows.map(r => r[regionCol]).filter(Boolean))] : ['All regions']
  const categories = categoryCol ? ['All categories', ...new Set(rows.map(r => r[categoryCol]).filter(Boolean))] : ['All categories']
  const filtered = rows.filter(r =>
    (region === 'All regions' || String(r[regionCol]) === region) &&
    (category === 'All categories' || String(r[categoryCol]) === category) &&
    (!query || Object.values(r).some(v => String(v).toLowerCase().includes(query.toLowerCase())))
  )

  const totalSales = salesCol ? filtered.reduce((s,r)=>s+cleanNum(r[salesCol]),0) : 0
  const totalProfit = profitCol ? filtered.reduce((s,r)=>s+cleanNum(r[profitCol]),0) : 0
  const totalOrders = orderCol ? filtered.reduce((s,r)=>s+cleanNum(r[orderCol]),0) : filtered.length
  const avg = filtered.length ? totalSales / filtered.length : 0
  const moneyLabel = salesCol ? titleCase(salesCol) : 'Numeric total'

  const trendData = useMemo(() => {
    if (!salesCol) return []
    const map = new Map()
    filtered.forEach((r,i) => {
      const raw = dateCol ? r[dateCol] : `Record ${i+1}`
      let label = String(raw)
      if (raw instanceof Date) label = raw.toLocaleDateString('en-US',{month:'short',year:'2-digit'})
      else if (/^\d{4}-\d{2}/.test(label)) {
        const d = new Date(label)
        if (!isNaN(d)) label = d.toLocaleDateString('en-US',{month:'short',year:'2-digit'})
      }
      map.set(label, (map.get(label)||0)+cleanNum(r[salesCol]))
    })
    return [...map].map(([name,value])=>({name,value})).slice(0,12)
  }, [filtered, salesCol, dateCol])

  const categoryData = useMemo(() => {
    if (!categoryCol || !salesCol) return []
    const map = new Map()
    filtered.forEach(r=>map.set(String(r[categoryCol]||'Other'),(map.get(String(r[categoryCol]||'Other'))||0)+cleanNum(r[salesCol])))
    return [...map].map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value).slice(0,6)
  }, [filtered, categoryCol, salesCol])

  const topData = useMemo(() => {
    const col = productCol || categoryCol
    if (!col || !salesCol) return []
    const map = new Map()
    filtered.forEach(r=>map.set(String(r[col]||'Other'),(map.get(String(r[col]||'Other'))||0)+cleanNum(r[salesCol])))
    return [...map].map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value).slice(0,5)
  }, [filtered, productCol, categoryCol, salesCol])

  const insights = useMemo(() => {
    const list = []
    if (salesCol && filtered.length) {
      const best = [...topData].sort((a,b)=>b.value-a.value)[0]
      if (best) list.push(`${best.name} is the top performer with ${fmt(best.value)} in ${moneyLabel.toLowerCase()}.`)
      if (profitCol && totalSales) list.push(`Overall profit margin is ${(totalProfit/totalSales*100).toFixed(1)}%.`)
      if (trendData.length > 1) {
        const first=trendData[0].value, last=trendData[trendData.length-1].value
        const change=first ? ((last-first)/first*100) : 0
        list.push(`The latest period is ${Math.abs(change).toFixed(1)}% ${change>=0?'above':'below'} the first period in this view.`)
      }
    }
    list.push(`${filtered.length} records are currently included in your dashboard.`)
    return list
  }, [topData, filtered, salesCol, profitCol, totalSales, totalProfit, trendData, moneyLabel])

  async function loadFiles(fileList) {
    const files = Array.from(fileList || [])
    if (!files.length) return
    setIsLoadingFiles(true)
    setNotice('')
    try {
      const datasets = await Promise.all(files.map(readUploadFile))
      const combined = combineDatasets(datasets)
      if (!combined.headers.length || !combined.rows.length) throw new Error('No usable data rows found in the selected files.')
      const names = datasets.map(dataset => dataset.fileName)
      setRows(combined.rows)
      setDataHeaders(combined.headers)
      setLoadedFileNames(names)
      setFileName(names.length === 1 ? names[0] : `${names.length} files combined`)
      setSheetNames(datasets.map(dataset => dataset.sheetName))
      setSheet(datasets[0].sheetName)
      setRegion('All regions')
      setCategory('All categories')
      setQuery('')
      setNotice(`Successfully combined ${fmt(combined.rows.length)} rows from ${fmt(names.length)} ${names.length === 1 ? 'file' : 'files'}.`)
    } catch (error) {
      setNotice(`Could not load the selected files: ${error.message}`)
    } finally {
      setIsLoadingFiles(false)
    }
  }

  function handleFileSelection(event) {
    const files = Array.from(event.target.files || [])
    event.target.value = ''
    loadFiles(files)
  }

  function changeSheet(name) {
    setSheet(name)
    // Workbook sheets are re-read from the selected file on upload in this simple client-side demo.
    setNotice(`Sheet selection changed to ${name}. Re-upload the workbook to load another sheet in this version.`)
  }
  function exportCsv() {
    const ws = XLSX.utils.json_to_sheet(filtered)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Dashboard Data')
    XLSX.writeFile(wb, 'datavision-filtered-data.xlsx')
  }
  function exportPdf() { window.print() }

  return (
    <div className={dark ? 'app dark' : 'app light'}>
      <aside className="sidebar">
        <div className="brand"><div className="brand-mark"><Activity size={21}/></div><div><b>DataVision<span> AI</span></b><small>SMART ANALYTICS</small></div></div>
        <div className="workspace"><div className="workspace-icon">DV</div><div><b>My Workspace</b><small>Personal account</small></div><ChevronDown size={15}/></div>
        <div className="nav-label">WORKSPACE</div>
        {[
          ['Overview',LayoutDashboard],['Charts',BarChart3],['Data Cleaning',Table2],['Data Explorer',Database],['AI Insights',Sparkles],['AI Chat',MessageCircle]
        ].map(([label,Icon])=><button key={label} className={`nav-item ${active===label?'selected':''}`} onClick={()=>setActive(label)}><Icon size={18}/>{label}{label==='AI Insights'&&<span className="new-tag">NEW</span>}</button>)}
        <div className="nav-label second">MANAGE</div>
        <button className="nav-item" onClick={()=>fileRef.current?.click()}><FolderOpen size={18}/>Upload data</button>
        <button className="nav-item" onClick={()=>setNotice('Settings are coming soon.')}><Settings2 size={18}/>Settings</button>
        <div className="sidebar-bottom"><div className="help-card"><div className="help-icon"><CircleHelp size={17}/></div><b>Need a hand?</b><p>Explore your data and turn it into decisions.</p><button onClick={()=>setNotice('Tip: Upload a spreadsheet with column headers in the first row.')}>View quick guide <ChevronRight size={14}/></button></div><div className="profile"><div className="avatar">TS</div><div><b>Workspace User</b><small>Free plan</small></div><MoreHorizontal size={18}/></div></div>
      </aside>

      <main className="main">
        <header className="topbar">
          <div className="breadcrumbs">Workspace <ChevronRight size={14}/> <b>{active}</b></div>
          <div className="top-actions"><div className="status"><span/> All systems operational</div><button className="icon-btn" title="Toggle theme" onClick={()=>setDark(!dark)}>{dark?<Sun size={18}/>:<Moon size={18}/>}</button><button className="icon-btn" title="Notifications" onClick={()=>setNotice('You are all caught up!')}><Bell size={18}/></button><div className="top-avatar">TS</div></div>
        </header>

        <div className="content">
          <div className="page-heading">
            <div><div className="eyebrow"><span className="eyebrow-dot"/> YOUR DATA, MADE CLEAR</div><h1>{active==='Overview'?'Analytics overview':active}</h1><p>Understand your data at a glance. Discover trends, patterns and opportunities.</p></div>
            <div className="heading-actions"><button className="btn secondary" onClick={exportCsv}><ArrowDownToLine size={16}/> Export data</button><button className="btn secondary" title="Print dashboard or save as PDF" aria-label="Print / Save PDF" onClick={exportPdf}><Printer size={16}/> Print</button><button className="btn primary" disabled={isLoadingFiles} onClick={()=>fileRef.current?.click()}><FileUp size={17}/>{isLoadingFiles?'Loading files...':'Upload files'}</button></div>
          </div>

          {notice && <div className="notice"><CheckCircle2 size={16}/><span>{notice}</span><button onClick={()=>setNotice('')}><X size={15}/></button></div>}

          <div className="file-strip"><div className="file-icon"><FileSpreadsheet size={19}/></div><div className="file-meta" title={loadedFileNames.join(', ')}><b>{fileName}</b><small>{rows.length.toLocaleString()} rows <span>•</span> {headers.length} columns <span>•</span> {loadedFileNames.length > 1 ? `${loadedFileNames.length} files combined` : 'Auto-analyzed'}</small></div><div className="file-badge"><span/> DATA READY</div><button className="text-btn" onClick={()=>fileRef.current?.click()}>Change files</button></div>

          <div className="section-head"><div><h2>Performance snapshot</h2><p>Key metrics calculated from your uploaded data</p></div><div className="filter-group"><Filter size={15}/><select value={region} onChange={e=>setRegion(e.target.value)}>{regions.map(x=><option key={x}>{x}</option>)}</select><select value={category} onChange={e=>setCategory(e.target.value)}>{categories.map(x=><option key={x}>{x}</option>)}</select></div></div>

          <div className="kpi-grid">
            <Kpi title={salesCol?titleCase(salesCol):'Numeric total'} value={salesCol?fmt(totalSales):fmt(filtered.length)} suffix={salesCol?'':' rows'} icon={TrendingUp} tone="purple" note="Total across filtered data"/>
            <Kpi title={profitCol?titleCase(profitCol):'Total records'} value={profitCol?fmt(totalProfit):fmt(filtered.length)} icon={Activity} tone="green" note={profitCol?'Calculated from your data':'Rows in current view'}/>
            <Kpi title={orderCol?titleCase(orderCol):'Data columns'} value={orderCol?fmt(totalOrders):fmt(headers.length)} icon={BarChart3} tone="blue" note={orderCol?'Sum of selected measure':'Detected columns'}/>
            <Kpi title="Average per record" value={salesCol?fmt(avg):fmt(filtered.length?filtered.length/Math.max(headers.length,1):0)} icon={Zap} tone="orange" note="Average value in current view"/>
          </div>

          {active==='Charts' ? <ChartGallery rows={filtered} headers={headers}/> :
          active==='Data Cleaning' ? <DataCleaning rows={rows} headers={headers} onApply={(nextRows,nextHeaders,message)=>{
            setRows(nextRows)
            setDataHeaders(nextHeaders)
            setRegion(current => !regionCol || nextRows.some(row => String(row[regionCol]) === current) ? current : 'All regions')
            setCategory(current => !categoryCol || nextRows.some(row => String(row[categoryCol]) === current) ? current : 'All categories')
            setQuery('')
            setNotice(message)
          }}/> :
          active==='AI Insights' ? <div className="insight-page"><div className="insight-title"><div className="ai-orb"><Sparkles size={24}/></div><div><h2>AI-powered data highlights</h2><p>Automatic observations based on your uploaded dataset.</p></div></div>{insights.map((item,i)=><div className="insight-row" key={i}><div className="insight-number">0{i+1}</div><p>{item}</p><CheckCircle2 size={18}/></div>)}<div className="disclaimer">These are rule-based observations generated in your browser, not advice from an external AI model.</div></div> :
          active==='AI Chat' ? <DataChat rows={filtered} headers={headers}/> :
          active==='Data Explorer' ? <DataTable rows={filtered} headers={headers} query={query} setQuery={setQuery}/> :
          <>
            <div className="charts-grid">
              <div className="panel trend-panel"><div className="panel-head"><div><h3>Performance trend</h3><p>{salesCol?`Trend by ${dateCol?titleCase(dateCol):'record'}`:'Upload numeric sales/revenue data to see a trend'}</p></div><span className="panel-menu"><MoreHorizontal size={19}/></span></div><div className="chart-area">{trendData.length?<ResponsiveContainer width="100%" height="100%"><AreaChart data={trendData} margin={{top:10,right:8,left:0,bottom:0}}><defs><linearGradient id="areaFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#8b7cff" stopOpacity={.35}/><stop offset="100%" stopColor="#8b7cff" stopOpacity={0}/></linearGradient></defs><CartesianGrid strokeDasharray="3 5" vertical={false} stroke="var(--grid)"/><XAxis dataKey="name" tick={{fill:'var(--muted)',fontSize:11}} axisLine={false} tickLine={false}/><YAxis tickFormatter={v=>v>=1000?`${Math.round(v/1000)}k`:v} tick={{fill:'var(--muted)',fontSize:11}} axisLine={false} tickLine={false} width={42}/><Tooltip contentStyle={{background:'var(--tooltip)',border:'1px solid var(--border)',borderRadius:12,color:'var(--text)'}} formatter={v=>fmt(v)}/><Area type="monotone" dataKey="value" stroke="#9888ff" strokeWidth={3} fill="url(#areaFill)" activeDot={{r:6,strokeWidth:0}}/></AreaChart></ResponsiveContainer>:<EmptyChart/>}</div></div>
              <div className="panel category-panel"><div className="panel-head"><div><h3>Category breakdown</h3><p>{categoryCol?`Distribution by ${titleCase(categoryCol)}`:'Categorical column not detected'}</p></div><span className="panel-menu"><MoreHorizontal size={19}/></span></div><div className="donut-wrap">{categoryData.length?<><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={categoryData} dataKey="value" nameKey="name" innerRadius={64} outerRadius={91} paddingAngle={4} stroke="none">{categoryData.map((x,i)=><Cell key={x.name} fill={COLORS[i%COLORS.length]}/>)}</Pie><Tooltip formatter={v=>fmt(v)} contentStyle={{background:'var(--tooltip)',border:'1px solid var(--border)',borderRadius:12}}/></PieChart></ResponsiveContainer><div className="donut-center"><b>{categoryData.length}</b><small>categories</small></div></>:<EmptyChart/>}</div><div className="legend">{categoryData.map((x,i)=><div key={x.name}><span style={{background:COLORS[i%COLORS.length]}}/>{x.name}<b>{fmt(x.value)}</b></div>)}</div></div>
            </div>
            <div className="bottom-grid">
              <div className="panel top-panel"><div className="panel-head"><div><h3>Top performers</h3><p>{productCol?titleCase(productCol):'Leading categories'} by {salesCol?titleCase(salesCol):'value'}</p></div><span className="subtle-label">TOP 5</span></div><div className="top-list">{topData.map((x,i)=><div className="top-row" key={x.name}><div className="rank">{String(i+1).padStart(2,'0')}</div><div className="top-info"><div className="top-label"><b>{x.name}</b><span>{fmt(x.value)}</span></div><div className="progress"><span style={{width:`${Math.max(4,x.value/(topData[0]?.value||1)*100)}%`,background:COLORS[i%COLORS.length]}}/></div></div></div>)}{!topData.length&&<EmptyChart/>}</div></div>
              <div className="panel insights-panel"><div className="panel-head"><div><h3><Sparkles size={16}/> Quick insights</h3><p>What your data is telling you</p></div><span className="ai-label">AI SUMMARY</span></div><div className="quick-insights">{insights.slice(0,3).map((x,i)=><div className="quick-item" key={i}><div className="quick-icon"><Lightbulb size={16}/></div><p>{x}</p></div>)}</div><button className="view-insights" onClick={()=>setActive('AI Insights')}>Explore all insights <ChevronRight size={15}/></button></div>
            </div>
          </>}
          <div className="footer"><span>© 2026 DataVision AI</span><span>Built to make data understandable <span className="footer-heart">✦</span></span><span>{filtered.length.toLocaleString()} records displayed</span></div>
        </div>
      </main>
      <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" multiple hidden onChange={handleFileSelection}/>
    </div>
  )
}

function DataCleaning({rows,headers,onApply}) {
  const [missingActions,setMissingActions] = useState({})
  const [typeActions,setTypeActions] = useState({})
  const [status,setStatus] = useState('')

  const duplicateCount = useMemo(() => getDuplicateRowCount(rows,headers),[rows,headers])
  const emptyColumns = useMemo(() => headers.filter(header =>
    rows.length > 0 && rows.every(row => isMissingValue(row[header]))
  ),[rows,headers])
  const missingColumns = useMemo(() => headers.map(header => ({
    header,
    count:rows.filter(row => isMissingValue(row[header])).length
  })).filter(column => column.count > 0),[rows,headers])
  const typeIssues = useMemo(() => headers.flatMap(header => {
    const values = rows.map(row => row[header]).filter(value => !isMissingValue(value))
    const numericCount = values.filter(isNumeric).length
    const invalidNumeric = values.filter(value => !isNumeric(value)).length
    const dateCount = values.filter(isDateValue).length
    const invalidDate = values.length-dateCount
    const numericColumnHint = /sales|revenue|amount|profit|price|cost|quantity|qty|order|count|total/i.test(header)
    const dateColumnHint = /date|time|month|year/i.test(header)
    if (numericCount > 0 && invalidNumeric > 0 && numericCount >= invalidNumeric) {
      return [{header,detected:`${numericCount} numeric, ${invalidNumeric} non-numeric`,suggested:'number',invalidCount:invalidNumeric}]
    }
    if ((dateCount > 0 && invalidDate > 0 && dateCount >= invalidDate) ||
      (dateColumnHint && dateCount < values.length)) {
      return [{header,detected:`${dateCount} date-like, ${invalidDate} other`,suggested:'date',invalidCount:invalidDate}]
    }
    if (numericColumnHint && numericCount === 0 && values.length) {
      return [{header,detected:`0 numeric, ${values.length} non-numeric`,suggested:'number',invalidCount:values.length}]
    }
    return []
  }),[rows,headers])

  function applyDuplicates() {
    const seen = new Set()
    const nextRows = rows.filter(row => {
      const key = JSON.stringify(headers.map(header => isMissingValue(row[header]) ? '' : String(row[header]).trim()))
      if (seen.has(key)) return false
      seen.add(key)
      return true
    })
    onApply(nextRows,headers,`Removed ${fmt(rows.length-nextRows.length)} duplicate rows.`)
    setStatus(`Removed ${fmt(rows.length-nextRows.length)} duplicate rows.`)
  }

  function applyEmptyColumns() {
    const nextHeaders = headers.filter(header => !emptyColumns.includes(header))
    const nextRows = rows.map(row => Object.fromEntries(nextHeaders.map(header => [header,row[header]])))
    onApply(nextRows,nextHeaders,`Removed ${fmt(emptyColumns.length)} empty columns.`)
    setStatus(`Removed ${fmt(emptyColumns.length)} empty columns.`)
  }

  function applyMissingValue(header) {
    const action = missingActions[header] || 'keep'
    if (action === 'keep') {
      setStatus(`No change made to ${titleCase(header)}.`)
      return
    }
    if (action === 'drop') {
      const nextRows = rows.filter(row => !isMissingValue(row[header]))
      const removed = rows.length-nextRows.length
      onApply(nextRows,headers,`Removed ${fmt(removed)} rows with missing ${titleCase(header)} values.`)
      setStatus(`Removed ${fmt(removed)} rows with missing ${titleCase(header)} values.`)
      return
    }
    const values = rows.map(row => row[header]).filter(value => !isMissingValue(value))
    let replacement
    if (action === 'zero') replacement = 0
    if (action === 'mode') {
      const counts = new Map()
      values.forEach(value => counts.set(String(value),(counts.get(String(value)) || 0)+1))
      replacement = values.find(value => counts.get(String(value)) === Math.max(...counts.values()))
    }
    if (action === 'median') {
      const sorted = values.filter(isNumeric).map(cleanNum).sort((a,b) => a-b)
      const middle = Math.floor(sorted.length/2)
      replacement = sorted.length
        ? (sorted.length%2 ? sorted[middle] : (sorted[middle-1]+sorted[middle])/2)
        : undefined
    }
    if (replacement === undefined) {
      setStatus(`Could not fill ${titleCase(header)}: no usable values were found.`)
      return
    }
    const nextRows = rows.map(row => isMissingValue(row[header]) ? {...row,[header]:replacement} : row)
    const filled = rows.filter(row => isMissingValue(row[header])).length
    onApply(nextRows,headers,`Filled ${fmt(filled)} missing ${titleCase(header)} values.`)
    setStatus(`Filled ${fmt(filled)} missing ${titleCase(header)} values.`)
  }

  function applyTypeConversion(header) {
    const target = typeActions[header] || typeIssues.find(issue => issue.header === header)?.suggested
    if (!target) return
    let converted = 0
    let invalid = 0
    const nextRows = rows.map(row => {
      const value = row[header]
      if (isMissingValue(value)) return row
      let nextValue = value
      if (target === 'number') {
        if (isNumeric(value)) {
          nextValue = cleanNum(value)
          converted += 1
        } else {
          nextValue = ''
          invalid += 1
        }
      } else if (target === 'date') {
        const normalized = normalizeDateValue(value)
        if (normalized) {
          nextValue = normalized
          converted += 1
        } else {
          nextValue = ''
          invalid += 1
        }
      } else {
        nextValue = String(value)
        converted += 1
      }
      return {...row,[header]:nextValue}
    })
    const message = `Converted ${fmt(converted)} ${titleCase(header)} values to ${target}${invalid ? `; ${fmt(invalid)} invalid values were left blank` : ''}.`
    onApply(nextRows,headers,message)
    setStatus(message)
  }

  const issueCount = duplicateCount + emptyColumns.length +
    missingColumns.reduce((sum,column) => sum+column.count,0) +
    typeIssues.reduce((sum,issue) => sum+issue.invalidCount,0)

  return <section className="cleaning-workspace" aria-label="Data cleaning tools">
    <div className="cleaning-intro"><div><h2>Data quality & cleaning</h2><p>Review detected issues before applying changes to your in-memory dataset.</p></div><div className={`quality-badge ${issueCount ? 'has-issues' : ''}`}><span/>{issueCount ? `${fmt(issueCount)} issues found` : 'Data looks clean'}</div></div>
    <div className="cleaning-summary">
      <div><span>Rows</span><strong>{fmt(rows.length)}</strong></div>
      <div><span>Columns</span><strong>{fmt(headers.length)}</strong></div>
      <div><span>Duplicate rows</span><strong>{fmt(duplicateCount)}</strong></div>
      <div><span>Missing cells</span><strong>{fmt(missingColumns.reduce((sum,column)=>sum+column.count,0))}</strong></div>
    </div>

    <article className="cleaning-card">
      <div className="cleaning-card-heading"><div><h3>Duplicate rows</h3><p>Exact matches across all current columns. The first occurrence is kept.</p></div><button className="btn primary" type="button" disabled={!duplicateCount} onClick={applyDuplicates}>Remove duplicates</button></div>
      <div className="cleaning-result">{duplicateCount ? `${fmt(duplicateCount)} duplicate ${duplicateCount===1?'row':'rows'} can be removed.` : 'No duplicate rows detected.'}</div>
    </article>

    <article className="cleaning-card">
      <div className="cleaning-card-heading"><div><h3>Empty columns</h3><p>Columns where every row has a blank value.</p></div><button className="btn primary" type="button" disabled={!emptyColumns.length} onClick={applyEmptyColumns}>Remove empty columns</button></div>
      <div className="cleaning-result">{emptyColumns.length ? emptyColumns.map(titleCase).join(', ') : 'No empty columns detected.'}</div>
    </article>

    <article className="cleaning-card">
      <div className="cleaning-card-heading"><div><h3>Missing values</h3><p>Choose how to handle blank cells in each column. Nothing changes until you apply it.</p></div></div>
      {missingColumns.length ? <div className="cleaning-table-wrap"><table className="cleaning-table"><thead><tr><th>Column</th><th>Missing</th><th>Handling</th><th/></tr></thead><tbody>
        {missingColumns.map(({header,count})=><tr key={header}><td>{titleCase(header)}</td><td>{fmt(count)} / {fmt(rows.length)}</td><td><select aria-label={`Missing value handling for ${titleCase(header)}`} value={missingActions[header] || 'keep'} onChange={event=>setMissingActions(current=>({...current,[header]:event.target.value}))}>
          <option value="keep">Keep as blank</option><option value="drop">Remove affected rows</option><option value="zero">Replace with 0</option><option value="mode">Fill with most common value</option><option value="median">Fill with numeric median</option>
        </select></td><td><button className="cleaning-apply" type="button" disabled={(missingActions[header] || 'keep')==='keep'} onClick={()=>applyMissingValue(header)}>Apply</button></td></tr>)}
      </tbody></table></div> : <div className="cleaning-result">No missing values detected.</div>}
    </article>

    <article className="cleaning-card">
      <div className="cleaning-card-heading"><div><h3>Possible data type issues</h3><p>Mixed or inconsistent values are flagged for review. Invalid values become blank during conversion.</p></div></div>
      {typeIssues.length ? <div className="cleaning-table-wrap"><table className="cleaning-table"><thead><tr><th>Column</th><th>Detected values</th><th>Convert to</th><th/></tr></thead><tbody>
        {typeIssues.map(issue=><tr key={issue.header}><td>{titleCase(issue.header)}</td><td>{issue.detected}</td><td><select aria-label={`Convert ${titleCase(issue.header)} data type`} value={typeActions[issue.header] || issue.suggested} onChange={event=>setTypeActions(current=>({...current,[issue.header]:event.target.value}))}><option value="number">Number</option><option value="date">Date</option><option value="text">Text</option></select></td><td><button className="cleaning-apply" type="button" onClick={()=>applyTypeConversion(issue.header)}>Convert</button></td></tr>)}
      </tbody></table></div> : <div className="cleaning-result">No mixed numeric or date-like columns detected.</div>}
    </article>
    {status && <p className="cleaning-status" role="status">{status}</p>}
    <p className="cleaning-disclaimer">Cleaning updates the current session only. Export the cleaned data to keep a copy.</p>
  </section>
}

function DataChat({rows,headers}) {
  const [messages, setMessages] = useState([
    { role:'assistant', text:'Hi! Ask me a question about the data in your dashboard. I can calculate totals, averages, and highest or lowest values by product, region, or category.' }
  ])
  const [draft, setDraft] = useState('')
  const endRef = useRef(null)

  useEffect(() => {
    endRef.current?.scrollIntoView({ behavior:'smooth', block:'end' })
  }, [messages])

  function sendMessage(question = draft) {
    const text = question.trim()
    if (!text) return
    const answer = answerDataQuestion(text, rows, headers)
    setMessages(current => [...current, { role:'user', text }, { role:'assistant', text:answer }])
    setDraft('')
  }

  return <section className="chat-panel" aria-label="Chat with your data">
    <div className="chat-heading">
      <div className="ai-orb"><MessageCircle size={22}/></div>
      <div><h2>AI Chat with Your Data</h2><p>Ask in plain English. Answers use {fmt(rows.length)} records currently in your dashboard view.</p></div>
    </div>
    <div className="chat-suggestions" aria-label="Suggested questions">
      {['What is my total sales?','Which product has the highest profit?','Which region has the lowest sales?'].map(question =>
        <button key={question} type="button" onClick={() => sendMessage(question)}>{question}</button>
      )}
    </div>
    <div className="chat-messages" role="log" aria-live="polite">
      {messages.map((message,index) => <div className={`chat-message ${message.role}`} key={index}>
        <span className="chat-avatar">{message.role === 'assistant' ? <Sparkles size={15}/> : <UserRound size={15}/>}</span>
        <p>{message.text}</p>
      </div>)}
      <div ref={endRef}/>
    </div>
    <form className="chat-input" onSubmit={event => { event.preventDefault(); sendMessage() }}>
      <input value={draft} onChange={event => setDraft(event.target.value)} placeholder="Ask a question about your data..." aria-label="Your question" />
      <button type="submit" disabled={!draft.trim()} aria-label="Send question"><Send size={17}/></button>
    </form>
    <p className="chat-disclaimer">Answers are calculated locally from your data and current dashboard filters. Your file is not sent to an AI service.</p>
  </section>
}

function ChartGallery({rows,headers}) {
  const numericCols = headers.filter(header => rows.some(row => isNumeric(row[header])))
  const categoricalCols = headers.filter(header => !numericCols.includes(header))
  const dateCol = findColumn(headers,['date','month','year'])
  const dimensions = categoricalCols.filter(column => column !== dateCol)
  const defaultMetric = findColumn(numericCols,['sales','revenue','amount','total']) || numericCols[0] || ''
  const [metric,setMetric] = useState('')
  const [scatterX,setScatterX] = useState('')
  const [scatterY,setScatterY] = useState('')
  const selectedMetric = numericCols.includes(metric) ? metric : defaultMetric
  const xMetric = numericCols.includes(scatterX) ? scatterX : numericCols[0]
  const yMetric = numericCols.includes(scatterY) ? scatterY : numericCols.find(column => column !== xMetric)
  const groupCol = dimensions.find(column => findColumn([column],['product','item'])) ||
    dimensions.find(column => findColumn([column],['category','segment','department'])) ||
    dimensions.find(column => findColumn([column],['region','state','city','country'])) ||
    dimensions[0]
  const heatRows = dimensions[0] || ''
  const heatColumns = dimensions[1] || ''

  const groupData = useMemo(() => {
    if (!selectedMetric || !groupCol) return []
    const values = new Map()
    rows.forEach(row => {
      const name = String(row[groupCol] ?? '').trim() || 'Unspecified'
      values.set(name,(values.get(name) || 0) + cleanNum(row[selectedMetric]))
    })
    return [...values].map(([name,value]) => ({name,value}))
      .sort((a,b) => b.value-a.value)
  }, [rows,selectedMetric,groupCol])

  const trendData = useMemo(() => {
    if (!selectedMetric) return []
    if (!dateCol) return rows.slice(0,20).map((row,index) => ({name:`Record ${index+1}`,value:cleanNum(row[selectedMetric]),sortKey:index}))
    const values = new Map()
    rows.forEach(row => {
      const raw = row[dateCol]
      const date = new Date(raw)
      const validDate = raw && !Number.isNaN(date.getTime())
      const key = validDate ? `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}` : String(raw || 'Unspecified')
      const name = validDate ? date.toLocaleDateString('en-US',{month:'short',year:'2-digit'}) : key
      const current = values.get(key) || {name,value:0,sortKey:validDate ? date.getTime() : Number.MAX_SAFE_INTEGER}
      current.value += cleanNum(row[selectedMetric])
      values.set(key,current)
    })
    return [...values.values()].sort((a,b) => a.sortKey-b.sortKey).slice(0,24)
  }, [rows,selectedMetric,dateCol])

  const scatterData = useMemo(() => rows.filter(row => isNumeric(row[xMetric]) && isNumeric(row[yMetric]))
    .map(row => ({x:cleanNum(row[xMetric]),y:cleanNum(row[yMetric])})), [rows,xMetric,yMetric])

  const heatmapData = useMemo(() => {
    if (!selectedMetric || !heatRows || !heatColumns) return {rows:[],columns:[],cells:new Map(),max:0}
    const rowNames = [...new Set(rows.map(row => String(row[heatRows] ?? '').trim() || 'Unspecified'))].slice(0,8)
    const columnNames = [...new Set(rows.map(row => String(row[heatColumns] ?? '').trim() || 'Unspecified'))].slice(0,8)
    const cells = new Map()
    rows.forEach(row => {
      const rowName = String(row[heatRows] ?? '').trim() || 'Unspecified'
      const columnName = String(row[heatColumns] ?? '').trim() || 'Unspecified'
      if (!rowNames.includes(rowName) || !columnNames.includes(columnName)) return
      const key = `${rowName}\u0000${columnName}`
      cells.set(key,(cells.get(key) || 0) + cleanNum(row[selectedMetric]))
    })
    return {rows:rowNames,columns:columnNames,cells,max:Math.max(...cells.values(),0)}
  }, [rows,selectedMetric,heatRows,heatColumns])

  const waterfallData = useMemo(() => {
    if (trendData.length < 2) return []
    const periods = trendData.slice(0,10)
    const steps = [{name:periods[0].name,offset:0,increase:periods[0].value,decrease:0}]
    let running = periods[0].value
    periods.slice(1).forEach(period => {
      const delta = period.value-running
      steps.push({
        name:period.name,
        offset:Math.min(running,period.value),
        increase:Math.max(delta,0),
        decrease:Math.max(-delta,0)
      })
      running = period.value
    })
    steps.push({name:'Latest',offset:0,increase:running,decrease:0})
    return steps
  }, [trendData])

  const charts = [
    {title:'Bar chart',description:`${groupCol ? titleCase(groupCol) : 'Categories'} ranked by ${selectedMetric ? titleCase(selectedMetric) : 'value'}`,icon:'▮',content:<div className="gallery-chart"><ChartEmpty hidden={Boolean(groupData.length)}/>{Boolean(groupData.length) && <ResponsiveContainer width="100%" height="100%"><BarChart data={groupData.slice(0,10)} margin={{top:8,right:8,left:0,bottom:34}}><CartesianGrid strokeDasharray="3 5" vertical={false} stroke="var(--grid)"/><XAxis dataKey="name" interval={0} angle={-25} textAnchor="end" height={55} tick={{fill:'var(--muted)',fontSize:10}} axisLine={false} tickLine={false}/><YAxis tickFormatter={value=>value>=1000?`${Math.round(value/1000)}k`:value} tick={{fill:'var(--muted)',fontSize:10}} axisLine={false} tickLine={false} width={42}/><Tooltip formatter={value=>fmt(value)} contentStyle={tooltipStyle}/><Bar dataKey="value" name={selectedMetric} fill="#8b7cff" radius={[5,5,0,0]} maxBarSize={42}/></BarChart></ResponsiveContainer>}</div>},
    {title:'Line chart',description:`${selectedMetric ? titleCase(selectedMetric) : 'Value'} over ${dateCol ? titleCase(dateCol) : 'record order'}`,icon:'⌁',content:<div className="gallery-chart"><ChartEmpty hidden={Boolean(trendData.length)}/>{Boolean(trendData.length) && <ResponsiveContainer width="100%" height="100%"><LineChart data={trendData} margin={{top:12,right:14,left:0,bottom:8}}><CartesianGrid strokeDasharray="3 5" vertical={false} stroke="var(--grid)"/><XAxis dataKey="name" tick={{fill:'var(--muted)',fontSize:10}} axisLine={false} tickLine={false}/><YAxis tickFormatter={value=>value>=1000?`${Math.round(value/1000)}k`:value} tick={{fill:'var(--muted)',fontSize:10}} axisLine={false} tickLine={false} width={42}/><Tooltip formatter={value=>fmt(value)} contentStyle={tooltipStyle}/><Line type="monotone" dataKey="value" name={selectedMetric} stroke="#55a7ff" strokeWidth={3} dot={{r:3,fill:'#55a7ff'}} activeDot={{r:6}}/></LineChart></ResponsiveContainer>}</div>},
    {title:'Pie chart',description:`Share of ${selectedMetric ? titleCase(selectedMetric) : 'value'} by ${groupCol ? titleCase(groupCol) : 'category'}`,icon:'◕',content:<div className="gallery-chart pie-gallery"><ChartEmpty hidden={Boolean(groupData.length)}/>{Boolean(groupData.length) && <ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={groupData.slice(0,8)} dataKey="value" nameKey="name" innerRadius={56} outerRadius={88} paddingAngle={3} stroke="none">{groupData.slice(0,8).map((item,index)=><Cell key={item.name} fill={COLORS[index%COLORS.length]}/>)}</Pie><Tooltip formatter={value=>fmt(value)} contentStyle={tooltipStyle}/><Legend verticalAlign="bottom" height={28} wrapperStyle={{fontSize:10,color:'var(--muted)'}}/></PieChart></ResponsiveContainer>}</div>},
    {title:'Scatter chart',description:`Relationship between ${xMetric ? titleCase(xMetric) : 'X'} and ${yMetric ? titleCase(yMetric) : 'Y'}`,icon:'⠿',content:<><div className="scatter-controls"><select aria-label="Scatter horizontal measure" value={xMetric || ''} onChange={event=>setScatterX(event.target.value)}>{numericCols.map(column=><option key={column} value={column}>{titleCase(column)}</option>)}</select><select aria-label="Scatter vertical measure" value={yMetric || ''} onChange={event=>setScatterY(event.target.value)}>{numericCols.filter(column=>column!==xMetric).map(column=><option key={column} value={column}>{titleCase(column)}</option>)}</select></div><div className="gallery-chart short-chart"><ChartEmpty hidden={Boolean(scatterData.length && yMetric)}/>{Boolean(scatterData.length && yMetric) && <ResponsiveContainer width="100%" height="100%"><ScatterChart margin={{top:10,right:14,left:0,bottom:8}}><CartesianGrid strokeDasharray="3 5" stroke="var(--grid)"/><XAxis type="number" dataKey="x" name={xMetric} tick={{fill:'var(--muted)',fontSize:10}} axisLine={false} tickLine={false}/><YAxis type="number" dataKey="y" name={yMetric} tick={{fill:'var(--muted)',fontSize:10}} axisLine={false} tickLine={false} width={42}/><ZAxis range={[45,45]}/><Tooltip cursor={{strokeDasharray:'3 3'}} formatter={value=>fmt(value)} contentStyle={tooltipStyle}/><Scatter data={scatterData} fill="#37c9a5" fillOpacity={0.75}/></ScatterChart></ResponsiveContainer>}</div></>},
    {title:'Treemap',description:`Relative ${selectedMetric ? titleCase(selectedMetric) : 'value'} by ${groupCol ? titleCase(groupCol) : 'category'}`,icon:'▦',content:<div className="gallery-chart"><ChartEmpty hidden={Boolean(groupData.length)}/>{Boolean(groupData.length) && <ResponsiveContainer width="100%" height="100%"><Treemap data={groupData.slice(0,20)} dataKey="value" aspectRatio={4/3} content={<TreemapTile/>}><Tooltip formatter={value=>fmt(value)} contentStyle={tooltipStyle}/></Treemap></ResponsiveContainer>}</div>},
    {title:'Heatmap',description:heatRows && heatColumns ? `${titleCase(heatRows)} vs ${titleCase(heatColumns)}` : 'Compare two categorical columns',icon:'▧',content:<HeatmapView data={heatmapData} metric={selectedMetric}/>},
    {title:'Waterfall chart',description:`Period-over-period ${selectedMetric ? titleCase(selectedMetric) : 'value'} movement`,icon:'▥',content:<div className="gallery-chart"><ChartEmpty hidden={Boolean(waterfallData.length)}/>{Boolean(waterfallData.length) && <ResponsiveContainer width="100%" height="100%"><BarChart data={waterfallData} margin={{top:10,right:8,left:0,bottom:8}}><CartesianGrid strokeDasharray="3 5" vertical={false} stroke="var(--grid)"/><XAxis dataKey="name" tick={{fill:'var(--muted)',fontSize:9}} axisLine={false} tickLine={false}/><YAxis tickFormatter={value=>value>=1000?`${Math.round(value/1000)}k`:value} tick={{fill:'var(--muted)',fontSize:10}} axisLine={false} tickLine={false} width={42}/><Tooltip formatter={value=>fmt(value)} contentStyle={tooltipStyle}/><Bar dataKey="offset" stackId="waterfall" fill="transparent" stroke="none" legendType="none" isAnimationActive={false}/><Bar dataKey="increase" name="Increase / total" stackId="waterfall" fill="#37c9a5" radius={[3,3,0,0]} isAnimationActive={false}/><Bar dataKey="decrease" name="Decrease" stackId="waterfall" fill="#ef7d9b" radius={[3,3,0,0]} isAnimationActive={false}/></BarChart></ResponsiveContainer>}</div>},
    {title:'Funnel chart',description:`${selectedMetric ? titleCase(selectedMetric) : 'Value'} by ${groupCol ? titleCase(groupCol) : 'stage'}, ranked`,icon:'▽',content:<div className="gallery-chart"><ChartEmpty hidden={Boolean(groupData.length)}/>{Boolean(groupData.length) && <ResponsiveContainer width="100%" height="100%"><FunnelChart><Tooltip formatter={value=>fmt(value)} contentStyle={tooltipStyle}/><Funnel dataKey="value" data={groupData.slice(0,8)} isAnimationActive={false}>{groupData.slice(0,8).map((item,index)=><Cell key={item.name} fill={COLORS[index%COLORS.length]}/>) }<LabelList position="right" fill="var(--text)" stroke="none" dataKey="name" fontSize={10}/></Funnel></FunnelChart></ResponsiveContainer>}</div>}
  ]

  return <section className="chart-gallery" aria-label="Professional chart gallery">
    <div className="gallery-toolbar"><div><h2>Professional charts</h2><p>Explore your filtered data in eight interactive chart formats.</p></div><label>Measure<select value={selectedMetric} onChange={event=>setMetric(event.target.value)} disabled={!numericCols.length}>{numericCols.map(column=><option key={column} value={column}>{titleCase(column)}</option>)}</select></label></div>
    {!numericCols.length && <div className="gallery-empty"><BarChart3 size={22}/>This dataset has no numeric columns to chart. Upload a file with a numeric measure.</div>}
    <div className="chart-gallery-grid">{charts.map((chart,index)=><article className="chart-card" key={chart.title}>
      <div className="chart-card-heading"><span className="chart-card-icon" style={{'--chart-accent':COLORS[index%COLORS.length]}}>{chart.icon}</span><div><h3>{chart.title}</h3><p>{chart.description}</p></div></div>
      {chart.content}
    </article>)}</div>
    <p className="gallery-note">Charts reflect the current dashboard filters. All calculations run locally in your browser.</p>
  </section>
}

const tooltipStyle = {background:'var(--tooltip)',border:'1px solid var(--border)',borderRadius:10,color:'var(--text)',fontSize:11}

function ChartEmpty({hidden}) {
  return hidden ? null : <div className="chart-empty">Add the required numeric and categorical columns to visualize this chart.</div>
}

function TreemapTile({x,y,width,height,name,value,index,depth}) {
  if (!width || !height || depth > 1) return null
  const color = COLORS[(index || 0)%COLORS.length]
  return <g>
    <rect x={x} y={y} width={width} height={height} rx={5} fill={color} fillOpacity={0.82} stroke="var(--panel)" strokeWidth={3}/>
    {width > 58 && height > 35 && <text x={x+9} y={y+19} fill="#fff" fontSize={11} fontWeight={600}>{String(name).length > 18 ? `${String(name).slice(0,16)}…` : name}</text>}
    {width > 72 && height > 52 && <text x={x+9} y={y+37} fill="#fff" fillOpacity={0.78} fontSize={10}>{fmt(value)}</text>}
  </g>
}

function HeatmapView({data,metric}) {
  if (!data.rows.length || !data.columns.length) return <div className="heatmap-empty">A heatmap needs two categorical columns and one numeric measure.</div>
  return <div className="heatmap-wrap"><div className="heatmap-grid" style={{'--heat-columns':data.columns.length}}>
    <span className="heatmap-corner">{titleCase(metric)}</span>
    {data.columns.map(column=><span className="heatmap-label column-label" key={column} title={column}>{column}</span>)}
    {data.rows.map(row=><React.Fragment key={row}>
      <span className="heatmap-label row-label" title={row}>{row}</span>
      {data.columns.map(column=>{
        const value = data.cells.get(`${row}\u0000${column}`) || 0
        const intensity = data.max ? 0.12 + value/data.max*0.78 : 0.12
        return <span className="heatmap-cell" key={`${row}-${column}`} title={`${row} · ${column}: ${fmt(value)}`} style={{background:`rgba(139,124,255,${intensity})`}}>{fmt(value)}</span>
      })}
    </React.Fragment>)}
  </div></div>
}

function Kpi({title,value,suffix='',icon:Icon,tone,note}) {
  return <div className="kpi-card"><div className="kpi-top"><span>{title}</span><div className={`kpi-icon ${tone}`}><Icon size={18}/></div></div><div className="kpi-value">{value}<small>{suffix}</small></div><div className="kpi-note"><span className="note-dot"/> {note}</div></div>
}
function EmptyChart(){return <div className="empty-chart"><BarChart3 size={28}/><span>Not enough matching data to visualize</span></div>}
function DataTable({rows,headers,query,setQuery}) {
  const shown=rows.filter(r=>!query||Object.values(r).some(v=>String(v).toLowerCase().includes(query.toLowerCase())))
  return <div className="panel table-panel"><div className="panel-head"><div><h3><Table2 size={17}/> Data explorer</h3><p>Browse and search your uploaded records</p></div><div className="search-box"><Search size={15}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search records..."/></div></div><div className="table-scroll"><table><thead><tr>{headers.map(h=><th key={h}>{titleCase(h)}</th>)}</tr></thead><tbody>{shown.slice(0,100).map((r,i)=><tr key={i}>{headers.map(h=><td key={h}>{String(r[h]??'')}</td>)}</tr>)}</tbody></table></div><div className="table-foot">Showing {Math.min(shown.length,100)} of {shown.length} matching records</div></div>
}

export default App
