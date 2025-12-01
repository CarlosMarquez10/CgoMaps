import express from 'express'
import dotenv from 'dotenv'
import path from 'path'
import { fileURLToPath } from 'url'
import { getClients } from './service/apis.js'
import ExcelJS from 'exceljs'

dotenv.config()
const app = express()
const Port = process.env.PORT || 3000

// Necesario para __dirname con ES modules
const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

// Configurar EJS como motor de vistas
app.set('view engine', 'ejs')
app.set('views', path.join(__dirname, 'views'))

// Servir archivos estáticos
app.use(express.static(path.join(__dirname, 'public')))
app.use(express.json())

// Configuración de Multer para subida de archivos
import multer from 'multer'
import fs from 'fs'
import xlsx from 'xlsx'

// Asegurar que existe el directorio de uploads
const uploadDir = path.join(__dirname, 'upload');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir);
}

// Directorio de proyectos/data
const dataDir = path.join(__dirname, 'data');
if (!fs.existsSync(dataDir)) {
  fs.mkdirSync(dataDir);
}

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, uploadDir)
  },
  filename: function (req, file, cb) {
    cb(null, Date.now() + '-' + file.originalname)
  }
})

const upload = multer({
  storage: storage,
  limits: { fileSize: 200 * 1024 * 1024 } // 200MB limit
})

app.get('/maps', (req, res) => {
  res.render('map')  // Llama a views/map.ejs
  getClients()
    .then(data => {
      console.log('Datos recibidos:', data);
      // Aquí puedes trabajar con los datos recibidos
    })
    .catch(error => {
      console.error('Error al obtener los datos:', error);
    });
})

app.post('/upload', upload.single('excelFile'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No se subió ningún archivo' });
    }

    // Leer el archivo Excel
    const workbook = xlsx.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);

    res.json({
      message: 'Archivo subido correctamente',
      filename: req.file.filename,
      data: data
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Error al procesar el archivo' });
  }
})

app.get('/files', (req, res) => {
  fs.readdir(uploadDir, (err, files) => {
    if (err) {
      return res.status(500).json({ error: 'Error al leer el directorio' });
    }
    const excelFiles = files.filter(file => file.endsWith('.xlsx') || file.endsWith('.xls'));
    res.json(excelFiles);
  });
})

app.get('/projects', (req, res) => {
  fs.readdir(dataDir, (err, files) => {
    if (err) {
      return res.status(500).json({ error: 'Error al leer el directorio' })
    }
    const excelFiles = files.filter(file => file.endsWith('.xlsx') || file.endsWith('.xls'))
    res.json(excelFiles)
  })
})

app.delete('/projects/:filename', (req, res) => {
  try {
    const raw = req.params.filename
    const safe = path.basename(raw)
    const fullPath = path.join(dataDir, safe)
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    fs.unlink(fullPath, (err) => {
      if (err) {
        return res.status(500).json({ error: 'No se pudo borrar el archivo' })
      }
      res.json({ message: 'Archivo borrado correctamente' })
    })
  } catch (e) {
    res.status(500).json({ error: 'Error al borrar el archivo' })
  }
})

app.get('/projects/:filename/points', (req, res) => {
  try {
    const raw = req.params.filename
    const safe = path.basename(raw)
    const fullPath = path.join(dataDir, safe)
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    const workbook = xlsx.readFile(fullPath, { cellDates: true })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    let rows = xlsx.utils.sheet_to_json(sheet)
    const descQuery = (req.query && typeof req.query.description === 'string') ? req.query.description.trim() : ''
    const descListQuery = (req.query && typeof req.query.descriptions === 'string') ? req.query.descriptions.trim() : ''
    if (descListQuery) {
      const list = descListQuery.split(',').map(s => s.trim().toUpperCase()).filter(Boolean)
      if (list.length) {
        rows = rows.filter(r => {
          const d = r.Description !== undefined && r.Description !== null ? String(r.Description).trim().toUpperCase() : ''
          return list.includes(d)
        })
      }
    } else if (descQuery) {
      const q = descQuery.toUpperCase()
      rows = rows.filter(r => {
        const d = r.Description !== undefined && r.Description !== null ? String(r.Description).trim().toUpperCase() : ''
        return d === q
      })
    }
    const defaultLat = 7.896031
    const defaultLng = -72.504365
    const getColor = (r) => {
      const corr = r['Color Correria']
      const motivo = r['Color Motivo']
      const tarea = r['Color Tarea']
      const corrColor = typeof corr === 'string' && corr.trim() ? corr.trim() : ''
      const motivoColor = typeof motivo === 'string' && motivo.trim() ? motivo.trim() : ''
      const tareaColor = typeof tarea === 'string' && tarea.trim() ? tarea.trim() : ''
      return corrColor || motivoColor || tareaColor || '#202020'
    }
    const parseGps = (v) => {
      if (v === undefined || v === null) return { lat: defaultLat, lng: defaultLng }
      const s = String(v).trim()
      if (!s || !s.includes(',')) return { lat: defaultLat, lng: defaultLng }
      const parts = s.split(',').map(p => parseFloat(p.trim()))
      if (parts.length < 2 || Number.isNaN(parts[0]) || Number.isNaN(parts[1])) {
        return { lat: defaultLat, lng: defaultLng }
      }
      let a = parts[0], b = parts[1]
      // Decide orden: lat [-90,90], lon [-180,180]
      if (Math.abs(a) > 90 && Math.abs(b) <= 90) {
        return { lat: b, lng: a }
      }
      if (Math.abs(b) > 90 && Math.abs(a) <= 90) {
        return { lat: a, lng: b }
      }
      return { lat: a, lng: b }
    }
    const toDateString = (d) => {
      const dd = String(d.getUTCDate()).padStart(2, '0')
      const mm = String(d.getUTCMonth() + 1).padStart(2, '0')
      const yyyy = d.getUTCFullYear()
      return `${dd}/${mm}/${yyyy}`
    }
    const excelSerialToDate = (n) => new Date(Math.round((n - 25569) * 86400 * 1000))
    const normalizeFecha = (v) => {
      if (v instanceof Date) return toDateString(v)
      if (typeof v === 'number') return toDateString(excelSerialToDate(v))
      const t = Date.parse(v)
      if (!Number.isNaN(t)) return toDateString(new Date(t))
      return String(v || '')
    }
    const wantedFields = ['Cliente','Ruta','Dirección','Nombre','Ciclo','Tarea','Id. Orden Trabajo','Revisión','Nombre Localidad','Gps','Fecha Solicitud','Description','EstadoOts','Tpl','Color Correria','Color Tarea','Color Motivo','Motivo']
    const points = rows.map(r => {
      const { lat, lng } = parseGps(r['Gps'])
      const props = {}
      wantedFields.forEach(f => {
        if (f === 'Gps') {
          props[f] = `${lat.toFixed(6)}, ${lng.toFixed(6)}`
        } else if (f.toLowerCase().includes('fecha')) {
          props[f] = normalizeFecha(r[f])
        } else {
          props[f] = r[f] !== undefined && r[f] !== null ? String(r[f]) : ''
        }
      })
      const hasInfo = ['Description','EstadoOts','Tpl','Color Correria'].some(k => {
        const v = r[k]
        return v !== undefined && v !== null && String(v).trim().length > 0
      })
      props['__asignado'] = hasInfo ? '1' : '0'
      return {
        lat,
        lng,
        color: getColor(r),
        props
      }
    })
    res.json({ points })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Error al leer puntos del proyecto' })
  }
})

app.get('/projects/:filename/correrias', (req, res) => {
  try {
    const raw = req.params.filename
    const safe = path.basename(raw)
    const fullPath = path.join(dataDir, safe)
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    const workbook = xlsx.readFile(fullPath, { cellDates: true })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = xlsx.utils.sheet_to_json(sheet)
    const groups = {}
    rows.forEach(r => {
      const desc = r.Description !== undefined && r.Description !== null ? String(r.Description).trim() : ''
      if (!desc) return
      const tpl = r.Tpl !== undefined && r.Tpl !== null ? String(r.Tpl).trim() : ''
      const estado = r.EstadoOts !== undefined && r.EstadoOts !== null ? String(r.EstadoOts).trim() : ''
      const color = r['Color Correria'] !== undefined && r['Color Correria'] !== null ? String(r['Color Correria']).trim() : ''
      const estadoCorreria = r.EstadoCorreria !== undefined && r.EstadoCorreria !== null ? String(r.EstadoCorreria).trim() : ''
      if (!groups[desc]) groups[desc] = { count: 0, tplCounts: {}, estadoCounts: {}, colorCounts: {}, estadoCorrCounts: {} }
      groups[desc].count++
      if (tpl) groups[desc].tplCounts[tpl] = (groups[desc].tplCounts[tpl] || 0) + 1
      if (estado) groups[desc].estadoCounts[estado] = (groups[desc].estadoCounts[estado] || 0) + 1
      if (color) groups[desc].colorCounts[color] = (groups[desc].colorCounts[color] || 0) + 1
      if (estadoCorreria) groups[desc].estadoCorrCounts[estadoCorreria] = (groups[desc].estadoCorrCounts[estadoCorreria] || 0) + 1
    })
    const pickMost = (obj) => {
      const entries = Object.entries(obj || {})
      if (!entries.length) return ''
      entries.sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'es', { sensitivity: 'base' }))
      return entries[0][0]
    }
    const correrias = Object.entries(groups).map(([desc, g]) => ({
      description: desc,
      count: g.count,
      tpl: pickMost(g.tplCounts),
      estado: pickMost(g.estadoCounts),
      color: pickMost(g.colorCounts),
      estadoCorreria: pickMost(g.estadoCorrCounts)
    })).sort((a, b) => a.description.localeCompare(b.description, 'es', { sensitivity: 'base' }))
    res.json({ correrias })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Error al leer correrías del proyecto' })
  }
})

// Resumen por Motivo con conteo y color de "Color Tarea"
app.get('/projects/:filename/motivo-summary', (req, res) => {
  try {
    const raw = req.params.filename
    const safe = path.basename(raw)
    const fullPath = path.join(dataDir, safe)
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    const workbook = xlsx.readFile(fullPath, { cellDates: true })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = xlsx.utils.sheet_to_json(sheet)
    const map = {}
    rows.forEach(r => {
      const motivo = r.Motivo !== undefined && r.Motivo !== null ? String(r.Motivo).trim() : ''
      if (!motivo) return
      const colorMotivo = r['Color Motivo'] !== undefined && r['Color Motivo'] !== null ? String(r['Color Motivo']).trim() : ''
      const colorTarea = r['Color Tarea'] !== undefined && r['Color Tarea'] !== null ? String(r['Color Tarea']).trim() : ''
      const estado = r.EstadoOts !== undefined && r.EstadoOts !== null ? String(r.EstadoOts).trim().toLowerCase() : ''
      const isProg = estado && !/no\s*programad/.test(estado) && /programad/.test(estado)
      if (!map[motivo]) map[motivo] = { count: 0, colorMotivoCounts: {}, colorTareaCounts: {}, programadas: 0 }
      map[motivo].count++
      if (isProg) map[motivo].programadas++
      if (colorMotivo) map[motivo].colorMotivoCounts[colorMotivo] = (map[motivo].colorMotivoCounts[colorMotivo] || 0) + 1
      if (colorTarea) map[motivo].colorTareaCounts[colorTarea] = (map[motivo].colorTareaCounts[colorTarea] || 0) + 1
    })
    const pickMost = (obj) => {
      const entries = Object.entries(obj || {})
      if (!entries.length) return ''
      entries.sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'es', { sensitivity: 'base' }))
      return entries[0][0]
    }
    const summary = Object.entries(map).map(([motivo, info]) => {
      const cm = pickMost(info.colorMotivoCounts)
      const ct = pickMost(info.colorTareaCounts)
      const color = cm || ct || ''
      return { motivo, count: info.count, programadas: info.programadas, color }
    }).sort((a, b) => b.count - a.count || a.motivo.localeCompare(b.motivo, 'es', { sensitivity: 'base' }))
    res.json({ summary })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Error al generar resumen por motivo' })
  }
})

app.get('/projects/:filename/download', (req, res) => {
  try {
    const raw = req.params.filename
    const safe = path.basename(raw)
    const fullPath = path.join(dataDir, safe)
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    res.download(fullPath, safe)
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Error al descargar el archivo' })
  }
})

app.get('/projects/:filename/download-sac', (req, res) => {
  try {
    const raw = req.params.filename
    const safe = path.basename(raw)
    const fullPath = path.join(dataDir, safe)
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    const workbook = xlsx.readFile(fullPath, { cellDates: true })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = xlsx.utils.sheet_to_json(sheet)
    const mapTareaToSac = (t) => {
      if (t === undefined || t === null) return ''
      const raw = String(t).trim()
      const norm = raw.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase()
      if (norm === 'DESVIACION SIGNIFICATIVA') return '01'
      if (norm === 'SUSPENSION') return '02'
      if (norm === 'VERIFICACION') return '02'
      if (norm === 'RECONEXION') return '03'
      if (norm === 'VISITA PERSUASIVA') return '06'
      return ''
    }
    const data = rows.map(r => ({
      'Description': r.Description !== undefined && r.Description !== null ? String(r.Description) : '',
      'Id. Orden Trabajo': r['Id. Orden Trabajo'] !== undefined && r['Id. Orden Trabajo'] !== null ? String(r['Id. Orden Trabajo']) : '',
      'TareaSac': r.TareaSac !== undefined && r.TareaSac !== null && String(r.TareaSac).trim() ? String(r.TareaSac) : mapTareaToSac(r['Tarea'])
    }))
    data.sort((a, b) => String(a['Description']).localeCompare(String(b['Description']), 'es', { sensitivity: 'base' }))
    const outSheet = xlsx.utils.json_to_sheet(data)
    const outWb = xlsx.utils.book_new()
    xlsx.utils.book_append_sheet(outWb, outSheet, 'SAC')
    const base = safe.replace(/\.xlsx?$/i, '')
    const outName = `${base}-SAC.xlsx`
    const buf = xlsx.write(outWb, { type: 'buffer', bookType: 'xlsx' })
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    res.setHeader('Content-Disposition', `attachment; filename="${outName}"`)
    res.send(buf)
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Error al generar archivo para SAC' })
  }
})

app.get('/projects/:filename/download-correria', async (req, res) => {
  try {
    const raw = req.params.filename
    const safe = path.basename(raw)
    const fullPath = path.join(dataDir, safe)
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    const workbook = xlsx.readFile(fullPath, { cellDates: true })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = xlsx.utils.sheet_to_json(sheet)
    const groups = {}
    rows.forEach(r => {
      const desc = r.Description !== undefined && r.Description !== null ? String(r.Description).trim() : ''
      if (!desc) return
      const id = r['Id. Orden Trabajo']
      const tpl = r.Tpl !== undefined && r.Tpl !== null ? String(r.Tpl).trim() : ''
      if (!groups[desc]) {
        groups[desc] = { count: 0, tplCounts: {} }
      }
      if (id !== undefined && id !== null && String(id).trim()) {
        groups[desc].count++
      } else {
        groups[desc].count++
      }
      if (tpl) {
        groups[desc].tplCounts[tpl] = (groups[desc].tplCounts[tpl] || 0) + 1
      }
    })
    const data = Object.entries(groups).map(([desc, info]) => {
      let terminal = ''
      const entries = Object.entries(info.tplCounts)
      if (entries.length) {
        entries.sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'es', { sensitivity: 'base' }))
        terminal = entries[0][0]
      }
      return { 'Descripción': desc, 'OTs': info.count, 'Terminal': terminal }
    })
    data.sort((a, b) => String(a['Descripción']).localeCompare(String(b['Descripción']), 'es', { sensitivity: 'base' }))
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('CORRERIA')
    ws.columns = [
      { header: 'Descripción', key: 'descripcion', width: 34 },
      { header: 'OTs', key: 'ots', width: 10 },
      { header: 'Terminal', key: 'terminal', width: 16 }
    ]
    data.forEach(r => {
      ws.addRow({ descripcion: r['Descripción'], ots: r['OTs'], terminal: r['Terminal'] })
    })
    const headerRow = ws.getRow(1)
    headerRow.height = 20
    headerRow.eachCell(c => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } }
      c.font = { bold: true, color: { argb: 'FFFFFFFF' } }
      c.alignment = { vertical: 'middle', horizontal: 'center' }
      c.border = {
        top: { style: 'thin', color: { argb: 'FFB0B0B0' } },
        left: { style: 'thin', color: { argb: 'FFB0B0B0' } },
        bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
        right: { style: 'thin', color: { argb: 'FFB0B0B0' } }
      }
    })
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r)
      row.eachCell(c => {
        c.border = {
          top: { style: 'thin', color: { argb: 'FFB0B0B0' } },
          left: { style: 'thin', color: { argb: 'FFB0B0B0' } },
          bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
          right: { style: 'thin', color: { argb: 'FFB0B0B0' } }
        }
        if (c.col === 2) {
          c.alignment = { horizontal: 'center', vertical: 'middle' }
        }
      })
    }
    const base = safe.replace(/\.xlsx?$/i, '')
    const outName = `${base}-Correria.xlsx`
    const buf = await wb.xlsx.writeBuffer()
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    res.setHeader('Content-Disposition', `attachment; filename="${outName}"`)
    res.send(Buffer.from(buf))
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Error al generar tabla de correría' })
  }
})

// Asegurar columnas Description y EstadoOts (y Color Correria, EstadoCorreria)
app.post('/projects/:filename/ensure-columns', (req, res) => {
  try {
    const raw = req.params.filename
    const safe = path.basename(raw)
    const fullPath = path.join(dataDir, safe)
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    const workbook = xlsx.readFile(fullPath, { cellDates: true })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = xlsx.utils.sheet_to_json(sheet)
    const ensured = rows.map(r => ({
      ...r,
      Description: r.Description !== undefined ? r.Description : '',
      EstadoOts: r.EstadoOts !== undefined ? r.EstadoOts : '',
      Tpl: r.Tpl !== undefined ? r.Tpl : '',
      'Color Correria': r['Color Correria'] !== undefined ? r['Color Correria'] : '',
      EstadoCorreria: r.EstadoCorreria !== undefined ? r.EstadoCorreria : ''
    }))
    const newSheet = xlsx.utils.json_to_sheet(ensured)
    workbook.Sheets[sheetName] = newSheet
    xlsx.writeFile(workbook, fullPath)
    res.json({ message: 'Columnas aseguradas', file: safe })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Error al asegurar columnas' })
  }
})

// Actualizar registro por Id. Orden Trabajo
app.post('/projects/:filename/update-record', (req, res) => {
  try {
    const raw = req.params.filename
    const { idOrdenTrabajo, description, estadoOts, color } = req.body || {}
    if (!idOrdenTrabajo) {
      return res.status(400).json({ error: 'idOrdenTrabajo es requerido' })
    }
    const safe = path.basename(raw)
    const fullPath = path.join(dataDir, safe)
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    const workbook = xlsx.readFile(fullPath, { cellDates: true })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = xlsx.utils.sheet_to_json(sheet)
    let updated = false
    const newRows = rows.map(r => {
      const id = r['Id. Orden Trabajo']
      if (String(id) === String(idOrdenTrabajo)) {
        updated = true
        return {
          ...r,
          Description: description !== undefined ? description : (r.Description || ''),
          EstadoOts: estadoOts !== undefined ? estadoOts : (r.EstadoOts || ''),
          'Color Correria': color !== undefined ? color : (r['Color Correria'] || '')
        }
      }
      return r
    })
    if (!updated) {
      return res.status(404).json({ error: 'Registro no encontrado' })
    }
    const newSheet = xlsx.utils.json_to_sheet(newRows)
    workbook.Sheets[sheetName] = newSheet
    xlsx.writeFile(workbook, fullPath)
    res.json({ message: 'Registro actualizado', file: safe })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Error al actualizar el registro' })
  }
})

app.post('/projects/:filename/unassign-record', (req, res) => {
  try {
    const raw = req.params.filename
    const { idOrdenTrabajo } = req.body || {}
    if (!idOrdenTrabajo) {
      return res.status(400).json({ error: 'idOrdenTrabajo es requerido' })
    }
    const safe = path.basename(raw)
    const fullPath = path.join(dataDir, safe)
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    const workbook = xlsx.readFile(fullPath, { cellDates: true })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = xlsx.utils.sheet_to_json(sheet)
    let updated = false
    const newRows = rows.map(r => {
      const id = r['Id. Orden Trabajo']
      if (String(id) === String(idOrdenTrabajo)) {
        updated = true
        return {
          ...r,
          Description: '',
          EstadoOts: '',
          'Color Correria': '',
          Tpl: ''
        }
      }
      return r
    })
    if (!updated) {
      return res.status(404).json({ error: 'Registro no encontrado' })
    }
    const newSheet = xlsx.utils.json_to_sheet(newRows)
    workbook.Sheets[sheetName] = newSheet
    xlsx.writeFile(workbook, fullPath)
    res.json({ message: 'Registro desprogramado', file: safe })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Error al desprogramar el registro' })
  }
})

app.post('/projects/merge', (req, res) => {
  try {
    const { target, source } = req.body || {}
    if (!target || !source) {
      return res.status(400).json({ error: 'target y source son requeridos' })
    }
    const safeTarget = path.basename(target)
    const safeSource = path.basename(source)
    const targetPath = path.join(dataDir, safeTarget)
    const sourcePath = path.join(dataDir, safeSource)
    if (![safeTarget, safeSource].every(n => n.endsWith('.xlsx') || n.endsWith('.xls'))) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(targetPath) || !fs.existsSync(sourcePath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    const tWb = xlsx.readFile(targetPath, { cellDates: true })
    const sWb = xlsx.readFile(sourcePath, { cellDates: true })
    const tSheetName = tWb.SheetNames[0]
    const sSheetName = sWb.SheetNames[0]
    const tSheet = tWb.Sheets[tSheetName]
    const sSheet = sWb.Sheets[sSheetName]
    const tRows = xlsx.utils.sheet_to_json(tSheet)
    const sRows = xlsx.utils.sheet_to_json(sSheet)
    const byId = {}
    tRows.forEach(r => {
      const id = r['Id. Orden Trabajo']
      if (id !== undefined && id !== null && String(id).trim()) byId[String(id).trim()] = r
    })
    let created = 0
    let updated = 0
    sRows.forEach(r => {
      const id = r['Id. Orden Trabajo']
      const key = id !== undefined && id !== null ? String(id).trim() : ''
      if (key && byId[key]) {
        const dst = byId[key]
        Object.keys(r).forEach(k => {
          const v = r[k]
          if (v !== undefined && v !== null && String(v).trim().length > 0) dst[k] = v
        })
        updated++
      } else {
        tRows.push(r)
        if (key) byId[key] = r
        created++
      }
    })
    const newSheet = xlsx.utils.json_to_sheet(tRows)
    tWb.Sheets[tSheetName] = newSheet
    xlsx.writeFile(tWb, targetPath)
    res.json({ message: 'Fusionado', created, updated, target: safeTarget })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Error al unificar proyectos' })
  }
})

// Actualizar Tpl por Description (bulk)
app.post('/projects/:filename/update-tpl-by-description', (req, res) => {
  try {
    const raw = req.params.filename
    const { description, tpl } = req.body || {}
    if (!description) {
      return res.status(400).json({ error: 'description es requerido' })
    }
    const safe = path.basename(raw)
    const fullPath = path.join(dataDir, safe)
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    const workbook = xlsx.readFile(fullPath, { cellDates: true })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = xlsx.utils.sheet_to_json(sheet)
    let updatedCount = 0
    const newRows = rows.map(r => {
      if ((r.Description || '') === description) {
        updatedCount++
        return { ...r, Tpl: tpl !== undefined ? tpl : (r.Tpl || ''), EstadoCorreria: 'guardado' }
      }
      return r
    })
    const newSheet = xlsx.utils.json_to_sheet(newRows)
    workbook.Sheets[sheetName] = newSheet
    xlsx.writeFile(workbook, fullPath)
    res.json({ message: 'Tpl actualizado por Description', updatedCount })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Error al actualizar Tpl por Description' })
  }
})
app.get('/files/:filename/columns', (req, res) => {
  try {
    const raw = req.params.filename
    const safe = path.basename(raw)
    const fullPath = path.join(uploadDir, safe)
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    const workbook = xlsx.readFile(fullPath)
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 })
    let headers = Array.isArray(rows) && rows.length > 0 ? rows[0] : []
    if (!headers || headers.length === 0) {
      const objRows = xlsx.utils.sheet_to_json(sheet)
      headers = objRows.length > 0 ? Object.keys(objRows[0]) : []
    }
    headers = headers.map(h => String(h)).filter(h => h && h.trim().length > 0)
    res.json({ columns: headers })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Error al leer columnas' })
  }
})

app.get('/files/:filename/column-summary', (req, res) => {
  try {
    const raw = req.params.filename
    const column = req.query.column
    if (!column) {
      return res.status(400).json({ error: 'Parámetro column es requerido' })
    }
    const safe = path.basename(raw)
    const fullPath = path.join(uploadDir, safe)
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    const workbook = xlsx.readFile(fullPath, { cellDates: true })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = xlsx.utils.sheet_to_json(sheet)
    const counts = {}
    const isDateColumn = /fecha|date/i.test(column)
    const toDateString = (d) => {
      const dd = String(d.getUTCDate()).padStart(2, '0')
      const mm = String(d.getUTCMonth() + 1).padStart(2, '0')
      const yyyy = d.getUTCFullYear()
      return `${dd}/${mm}/${yyyy}`
    }
    const excelSerialToDate = (n) => new Date(Math.round((n - 25569) * 86400 * 1000))
    const normalizeValue = (v) => {
      if (!isDateColumn) return v
      if (v instanceof Date) return toDateString(v)
      if (typeof v === 'number') return toDateString(excelSerialToDate(v))
      const t = Date.parse(v)
      if (!Number.isNaN(t)) return toDateString(new Date(t))
      return v
    }
    rows.forEach(r => {
      const val = normalizeValue(r[column])
      if (val === undefined || val === null) return
      const key = String(val).trim()
      if (!key) return
      counts[key] = (counts[key] || 0) + 1
    })
    const summary = Object.entries(counts).map(([value, count]) => ({ value, count }))
      .sort((a, b) => b.count - a.count || a.value.localeCompare(b.value))
    res.json({ column, summary })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Error al generar resumen' })
  }
})


app.delete('/files/:filename', (req, res) => {
  try {
    const raw = req.params.filename
    const safe = path.basename(raw)
    const fullPath = path.join(uploadDir, safe)
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    fs.unlink(fullPath, (err) => {
      if (err) {
        return res.status(500).json({ error: 'No se pudo borrar el archivo' })
      }
      res.json({ message: 'Archivo borrado correctamente' })
    })
  } catch (e) {
    res.status(500).json({ error: 'Error al borrar el archivo' })
  }
})


app.post('/files/:filename/colorized', (req, res) => {
  try {
    const raw = req.params.filename
    const safe = path.basename(raw)
    const fullPath = path.join(uploadDir, safe)
    const { column, colors } = req.body || {}
    if (!column || !colors || typeof colors !== 'object') {
      return res.status(400).json({ error: 'column y colors son requeridos' })
    }
    if (!safe.endsWith('.xlsx') && !safe.endsWith('.xls')) {
      return res.status(400).json({ error: 'Nombre de archivo inválido' })
    }
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Archivo no encontrado' })
    }
    const workbook = xlsx.readFile(fullPath, { cellDates: true })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = xlsx.utils.sheet_to_json(sheet)
    const colorColName = `Color ${column}`
    const mapTareaToSac = (t) => {
      if (t === undefined || t === null) return ''
      const raw = String(t).trim()
      const norm = raw.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase()
      if (norm === 'DESVIACION SIGNIFICATIVA') return '01'
      if (norm === 'SUSPENSION') return '02'
      if (norm === 'VERIFICACION') return '02'
      if (norm === 'RECONEXION') return '03'
      if (norm === 'VISITA PERSUASIVA') return '06'
      return ''
    }
    const getColor = (val) => {
      if (val === undefined || val === null) return '#000000'
      const key = String(val).trim()
      return colors[key] || '#000000'
    }
    const newRows = rows.map(r => ({ ...r, [colorColName]: getColor(r[column]), TareaSac: mapTareaToSac(r['Tarea']), EstadoCorreria: r.EstadoCorreria !== undefined ? r.EstadoCorreria : '' }))
    const newSheet = xlsx.utils.json_to_sheet(newRows)
    const newWb = xlsx.utils.book_new()
    xlsx.utils.book_append_sheet(newWb, newSheet, sheetName)
    const base = safe.replace(/\.xlsx?$/i, '')
    const outName = `${Date.now()}-${base}-colored.xlsx`
    const outPath = path.join(dataDir, outName)
    xlsx.writeFile(newWb, outPath)
    res.json({ message: 'Proyecto guardado', output: outName })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Error al guardar el proyecto' })
  }
})



app.listen(Port, () => {
  console.log(`Server listing on the port ${Port}`)
})

