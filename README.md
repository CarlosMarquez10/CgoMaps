# CGO Maps v3.0 🗺️

Una aplicación web moderna para la creación y gestión de rutas interactivas, similar a Google Earth, desarrollada con Node.js, Express y Leaflet.

## 🚀 Características

- **Mapas Interactivos**: Visualización de mapas utilizando Leaflet con múltiples capas y estilos
- **Creación de Rutas**: Herramientas intuitivas para crear y editar rutas personalizadas
- **Interfaz Moderna**: Diseño responsive con menús deslizantes y efectos visuales
- **Gestión de Clientes**: Integración con APIs para manejo de datos de clientes
- **Múltiples Vistas**: Soporte para diferentes tipos de mapas y visualizaciones

## 🛠️ Tecnologías Utilizadas

- **Backend**: Node.js con Express
- **Frontend**: HTML5, CSS3, JavaScript ES6+
- **Motor de Vistas**: EJS
- **Mapas**: Leaflet.js
- **Gestión de Variables**: dotenv
- **Desarrollo**: Nodemon para hot-reload

## 📋 Requisitos Previos

- Node.js (versión 14 o superior)
- npm o yarn
- Navegador web moderno

## 🔧 Instalación

1. Clona el repositorio:
```bash
git clone https://github.com/CarlosMarquez10/CgoMaps.git
cd CgoMaps
```

2. Instala las dependencias:
```bash
npm install
```

3. Configura las variables de entorno:
```bash
cp .env.example .env
# Edita el archivo .env con tus configuraciones
```

4. Inicia el servidor de desarrollo:
```bash
npm run dev
```

5. Abre tu navegador y visita:
```
http://localhost:3000/maps
```

## 🎯 Uso

### Navegación Básica
- Utiliza el botón de toggle en la esquina inferior izquierda para acceder al menú principal
- Los botones del menú permiten acceder a diferentes funcionalidades de la aplicación

### Creación de Rutas
- Selecciona la herramienta de creación de rutas desde el menú
- Haz clic en el mapa para agregar puntos de ruta
- Utiliza las opciones de edición para modificar rutas existentes

### Gestión de Datos
- La aplicación se conecta automáticamente con las APIs configuradas
- Los datos de clientes se cargan dinámicamente en el mapa

## 📁 Estructura del Proyecto

```
CGO-MAPS/
├── app.js                 # Servidor principal de Express
├── package.json           # Dependencias y scripts
├── .env                   # Variables de entorno (no incluido en git)
├── public/                # Archivos estáticos
│   └── leaflet/          # Recursos de Leaflet
│       ├── images/       # Iconos y sprites
│       ├── leaflet.css   # Estilos de Leaflet
│       └── leaflet.js    # Biblioteca Leaflet
├── service/              # Servicios y APIs
│   └── apis.js          # Configuración de APIs
└── views/               # Plantillas EJS
    └── map.ejs         # Vista principal del mapa
```

## 🔧 Scripts Disponibles

- `npm run dev`: Inicia el servidor en modo desarrollo con nodemon
- `npm test`: Ejecuta las pruebas (pendiente de implementación)

## 🌐 API Endpoints

- `GET /maps`: Renderiza la vista principal del mapa con funcionalidades completas

## 🎨 Personalización

### Estilos
Los estilos están integrados en el archivo `map.ejs` y pueden ser modificados para personalizar:
- Colores del tema
- Efectos de transparencia y blur
- Animaciones y transiciones
- Layout responsive

### Funcionalidades
Puedes extender las funcionalidades modificando:
- `service/apis.js`: Para agregar nuevas integraciones de API
- `views/map.ejs`: Para modificar la interfaz de usuario
- `app.js`: Para agregar nuevas rutas y middleware

## 🤝 Contribución

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📝 Licencia

Este proyecto está bajo la Licencia ISC. Ver el archivo `LICENSE` para más detalles.

## 👨‍💻 Autor

**Carlos Márquez**
- GitHub: [@CarlosMarquez10](https://github.com/CarlosMarquez10)
- Repositorio: [CgoMaps](https://github.com/CarlosMarquez10/CgoMaps.git)

## 🙏 Agradecimientos

- [Leaflet](https://leafletjs.com/) por la excelente biblioteca de mapas
- [Express.js](https://expressjs.com/) por el framework web
- [OpenStreetMap](https://www.openstreetmap.org/) por los datos de mapas

---

⭐ ¡No olvides dar una estrella al proyecto si te ha sido útil!