import express from 'express'
import dotenv from 'dotenv'
import path from 'path'
import { fileURLToPath } from 'url'
import { getClients } from './service/apis.js'

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
      return res.status(500).json({ error: 'Error al leer el directorio' });
    }
    const excelFiles = files.filter(file => file.endsWith('.xlsx') || file.endsWith('.xls'));
    res.json(excelFiles);
  });
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
    const rows = xlsx.utils.sheet_to_json(sheet)
    const defaultLat = 7.896031
    const defaultLng = -72.504365
    const cols = rows.length ? Object.keys(rows[0]) : []
    const colorCol = cols.find(c => c.toLowerCase().startsWith('color ')) || cols.find(c => c.toLowerCase().includes('color tarea'))
    const getColor = (r) => {
      const v = colorCol ? r[colorCol] : null
      return typeof v === 'string' && v.trim() ? v.trim() : '#28a745'
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
    const wantedFields = ['Cliente','Ruta','Dirección','Nombre','Ciclo','Tarea','Id. Orden Trabajo','Revisión','Nombre Localidad','Gps','Fecha Solicitud']
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

// Asegurar columnas Description y EstadoOts (y Color Correria)
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
      'Color Correria': r['Color Correria'] !== undefined ? r['Color Correria'] : ''
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
        return { ...r, Tpl: tpl !== undefined ? tpl : (r.Tpl || '') }
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
    const getColor = (val) => {
      if (val === undefined || val === null) return '#000000'
      const key = String(val).trim()
      return colors[key] || '#000000'
    }
    const newRows = rows.map(r => ({ ...r, [colorColName]: getColor(r[column]) }))
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

