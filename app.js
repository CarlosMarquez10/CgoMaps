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

app.get('/maps', (req, res) => {
    res.render('map')  // Llama a views/map.ejs
    getClients(clientes)
      .then(data => {
        console.log('Datos recibidos:', data);
        // Aquí puedes trabajar con los datos recibidos
      })
      .catch(error => {
        console.error('Error al obtener los datos:', error);
      });
})



app.listen(Port, () => {
    console.log(`Server listing on the port ${Port}`)
})

