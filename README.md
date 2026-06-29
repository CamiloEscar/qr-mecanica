# QR para Talleres Mecánicos

Sistema de gestión de historial de mantenimiento vehicular mediante códigos QR. Cada vehículo tiene un QR único que el cliente escanea con su teléfono para acceder al historial completo de servicios realizados en el taller.

**Stack**: Google Apps Script · Google Sheets · Drive · HTML/CSS/JS vanilla

---

## Demo en vivo

| Vista | URL |
|---|---|
| **Cliente** (escaneando QR) | https://script.google.com/macros/s/AKfycbw0pd6P9iqdUI8cvIIFjtPZfXNOekSIXxY2cYOCtA-DjPGSX50wro8dHToZ8HMi0dI0/exec?t=YBKhqkFhldCYMryK |
| **Panel del mecánico** | https://script.google.com/macros/s/AKfycbw0pd6P9iqdUI8cvIIFjtPZfXNOekSIXxY2cYOCtA-DjPGSX50wro8dHToZ8HMi0dI0/exec?accion=admin |
| **Diagnóstico del sistema** | https://script.google.com/macros/s/AKfycbw0pd6P9iqdUI8cvIIFjtPZfXNOekSIXxY2cYOCtA-DjPGSX50wro8dHToZ8HMi0dI0/exec?accion=diag |

---

## Funcionalidades

### Para el taller

- **Registrar clientes**: nombre, teléfono, email opcional, consentimiento
- **Registrar vehículos**: patente, marca, modelo, año, combustible
- **Cargar servicios**: fecha, kilometraje, descripción, repuestos, próximo mantenimiento, observaciones, fotos
- **Generar QRs automáticamente**: cada vehículo tiene un token único de 16 caracteres
- **Generar tarjeta A6 imprimible** con el QR + datos del vehículo + branding del taller
- **Descargar PDF A6** del QR para imprimir o mandar por email
- **Compartir por WhatsApp/email** desde el celular
- **Panel admin dark mode** responsive mobile-first
- **Gestión de sesión** con tokens temporales (24hs TTL)
- **Subir fotos por drag & drop** desde el celular

### Para el cliente

- **Escanear QR** con la cámara del teléfono (no requiere instalar app)
- **Ver historial cronológico** completo de servicios
- **Ver fotos** adjuntas a cada service (lightbox)
- **Ver próximo mantenimiento recomendado**
- **Contactar al taller** por WhatsApp o Instagram con un click

---

## Privacidad y seguridad

- El QR **NO apunta a la patente** del vehículo. Apunta a un **token aleatorio de 16 caracteres** (62^16 = 4.77 × 10^28 combinaciones posibles). Esto hace imposible que alguien adivine URLs de otros vehículos.
- La patente queda como dato **interno del taller**, nunca se expone públicamente.
- Los datos sensibles (admin key, mecánico password) se guardan en **Script Properties**, nunca en el código fuente.
- El cliente solo tiene acceso de **lectura**, nunca puede modificar datos.

---

## Arquitectura

```
┌─────────────────────────────────────────────┐
│ TALLER (admin)                              │
│ ┌──────────┐  ┌──────────────┐  ┌────────┐ │
│ │  Google  │◄►│ Apps Script  │◄►│ Drive  │ │
│ │  Sheet   │  │ Web App      │  │ (fotos)│ │
│ │ (4 tabs) │  │ (doGet/doPost│  │        │ │
│ └──────────┘  └──────┬───────┘  └────────┘ │
└──────────────────────┼─────────────────────┘
                       │ HTTPS
                       ▼
┌─────────────────────────────────────────────┐
│ CLIENTE (smartphone, sin app)               │
│ Escanea QR → URL con token → Apps Script   │
│ sirve HTML responsive con datos del vehículo│
└─────────────────────────────────────────────┘
```

**Decisiones técnicas clave:**

- **Google Sheets como DB**: gratis, persistente, fácil de administrar
- **Apps Script Web App** como backend + hosting: sin servidor que mantener
- **HTML Service** para servir las páginas (con todas sus limitaciones conocidas)
- **qr-code-styling** (cliente-side) para generar QRs visualmente bonitos con logo
- **jsPDF + html2canvas** para generar PDFs A6 descargables
- **Drag & drop HTML5** para subir fotos desde el celular
- **Sin frameworks JS** (sin React/Vue/Angular) para minimizar bundle size y complejidad

---

## Estructura del proyecto

```
qr-mecanica/
├── apps-script/
│   ├── Codigo.gs              # Backend Apps Script (1300+ líneas)
│   ├── ClientePage.html       # Vista del cliente al escanear QR
│   ├── AdminPage.html         # Panel admin para el mecánico
│   ├── LoginPage.html         # Pantalla de login
│   ├── QRPrintPage.html       # Tarjeta A6 imprimible + PDF
│   └── appsscript.json        # Manifest del proyecto Apps Script
├── templates/                 # Templates CSV para las hojas
│   ├── vehiculos.csv
│   ├── clientes.csv
│   ├── servicios.csv
│   └── accesos_log.csv
├── docs/
│   └── arquitectura.md        # Doc técnica detallada
├── README.md                  # Este archivo
└── .clasp.json                # Configuración de clasp CLI
```

---

## Hojas de Google Sheets

| Hoja | Columnas | Propósito |
|---|---|---|
| `Clientes` | ID, Nombre, Teléfono, Email, Fecha_Alta, Consentimiento | Datos de clientes |
| `Vehiculos` | ID, Token, Fecha_Alta, Activo, Patente, Marca, Modelo, Año, Combustible, ID_Cliente, URL_QR | Vehículos + tokens únicos |
| `Servicios` | ID, Fecha, ID_Vehiculo, Kilometraje, Descripcion, Repuestos, Proximo_Mantenimiento, Observaciones, Fotos_IDs, ID_Mecanico | Historial de servicios |
| `AccesosLog` | Timestamp, Token_Suffix, Email_O_Anonimo | Auditoría de accesos |

---

## Tech stack

| Componente | Tecnología | Costo |
|---|---|---|
| Base de datos | Google Sheets | $0 |
| Backend | Google Apps Script | $0 |
| Hosting web | Apps Script Web App | $0 |
| Almacenamiento fotos | Google Drive | $0 (hasta 15 GB) |
| Generación QR visual | qr-code-styling (cliente-side) | $0 |
| Generación PDF | jsPDF + html2canvas | $0 |
| Autenticación | Session tokens + Script Properties | $0 |

**Total: $0** para volúmenes de hasta ~2000 vehículos.

---

## Setup local (para clonar y desarrollar)

### Requisitos
- Cuenta de Google
- Node.js 18+
- [clasp](https://github.com/google/clasp) CLI: `npm install -g @google/clasp`

### Instalación

```bash
git clone https://github.com/CamiloEscar/qr-mecanica.git
cd qr-mecanica

# Login con Google
clasp login

# Crear spreadsheet y proyecto Apps Script
clasp create --title "QR Mecanica" --type sheets --rootDir ./apps-script

# Subir archivos
clasp push

# Abrir el editor y ejecutar setupInicial() una vez
clasp open-script

# Deploy como Web App
clasp deploy
```

Después del primer deploy, configurar el Web App manualmente desde el editor (Deploy → New deployment → Web app → Anyone) para que la URL sea accesible públicamente.

---

## Aprendizajes del proyecto

Este proyecto enseñó varias lecciones valiosas sobre Google Apps Script:

1. **`google.script.run` chain pattern**: SIEMPRE encadenar handlers ANTES de acceder a la función:
   ```javascript
   google.script.run
     .withSuccessHandler(...)
     .withFailureHandler(...)
     [functionName].apply(null, args);
   ```
   Si capturás la referencia antes de setear handlers, el await cuelga para siempre.

2. **Date serialization**: `google.script.run` no puede serializar objetos `Date`. Convertir a string ISO antes de retornar.

3. **CORS desde iframe sandbox**: Apps Script NO devuelve headers CORS. NO usar `fetch()`, usar siempre `google.script.run`.

4. **`window.location` en iframe**: apunta al sandbox, no al deploy. Usar template `<?= webappUrl ?>` inyectada desde backend.

5. **Múltiples deploys confunden el sandbox**: `clasp undeploy --all` antes de re-deployar.

---

## Roadmap futuro

- [ ] Recordatorios automáticos por WhatsApp (próximo mantenimiento)
- [ ] Recordatorios por email
- [ ] Presupuestos digitales con firma del cliente
- [ ] Historial multi-taller
- [ ] Panel admin con estadísticas (Looker Studio)
- [ ] Custom domain (Vercel proxy + Apps Script)
- [ ] App mobile nativa con React Native

---

## Licencia

MIT

---

## Autor

**Mecánica Martínez** - Taller mecánico automotor
Instagram: [@martinez_mecanicaautomotriz](https://www.instagram.com/martinez_mecanicaautomotriz/)