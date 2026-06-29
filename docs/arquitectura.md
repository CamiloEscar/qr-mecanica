# Arquitectura del sistema QR Mecánica

Documento de arquitectura técnica del sistema de historial vehicular por QR. Si estás por deployar, andá directo al [README.md](../README.md) — este doc es para entender el "por qué" de las decisiones.

## Stack tecnológico

| Capa | Tecnología | Por qué |
|---|---|---|
| Base de datos | Google Sheets | Gratis, accesible desde cualquier dispositivo, el taller ya sabe usarlo. Sin servidor que mantener. |
| Backend + serving | Apps Script Web App | Gratis, corre en infraestructura de Google, deploy en un clic. Maneja GET (página cliente) y POST (admin). |
| Hosting HTML | Apps Script `HtmlService` | Sin servidor de archivos. La plantilla `ClientePage.html` se inyecta con datos del servidor. |
| Almacenamiento fotos | Google Drive (carpeta `QR-Mecanica-Fotos`) | 15 GB gratis. Links públicos vía `setSharing(ANYONE_WITH_LINK)`. |
| Generación de QR | QuickChart API (`quickchart.io/qr`) | Servicio público, sin autenticación, sin rate limit conocido. Devuelve PNG listo para imprimir. |
| Front cliente | HTML + CSS inline (en `ClientePage.html`) | Cero dependencias externas. Carga instantánea. Mobile-first. |
| Triggers admin | Google Forms + Apps Script `onFormSubmit` | El personal del taller no necesita saber programar — completa Forms y el sistema hace el resto. |

**Costo total**: $0/mes hasta volúmenes medios. **Mantenimiento**: prácticamente cero, Google se encarga de la infra.

## Por qué token aleatorio de 16 caracteres

La pregunta clave de seguridad es: **¿qué URL se codifica en el QR?**

### Opción descartada: usar la patente

```
https://script.google.com/macros/s/.../exec?t=AB123CD
```

Problemas:
- **Las patentes son enumerable**: en Argentina hay ~15 millones de vehículos registrados. Un atacante con un script puede iterar `AB123CD`, `AB124CD`, etc.
- **Filtran PII**: la patente es dato personal identificable (ley 25.326). Ponerla en una URL accesible es regalarle info a cualquiera que vea el QR o intercepte el link.
- **Formato predecible**: 2 letras + 3 números + 2 letras (Mercosur) o AAA999 (viejo). Son patrones, no secretos.

### Opción elegida: token aleatorio de 16 caracteres alfanuméricos

```
https://script.google.com/macros/s/.../exec?t=a7f9k2m8x4p1q6r3
```

### La matemática

- Alfabeto: `a-z` (26) + `A-Z` (26) + `0-9` (10) = **62 caracteres**
- Largo: **16 caracteres**
- Espacio total: `62^16 = 4.77 × 10^28` combinaciones posibles

Para ponerlo en perspectiva:

| Escala | Cantidad |
|---|---|
| Combinaciones posibles | 4.770.000.000.000.000.000.000.000.000.000 |
| Vehículos registrados en Argentina | ~15.000.000 |
| Probabilidad de colisión en un taller de 1.000 autos | ~1 en `4.77 × 10^22` |
| Intentos por segundo para adivinar un token | incluso 1 millón/s → `1.5 × 10^15` años para cubrir el espacio |

**Conclusión**: un atacante no puede enumerar ni adivinar. Si ves un QR pegado en un auto y querés ver su historial, ya tenés el acceso (es el dueño o alguien de confianza). Si no, el token es criptográficamente aleatorio e impredecible.

### Por qué `Math.random()` es suficiente

`Math.random()` NO es criptográficamente seguro (CSPRNG). Para un token de sesión de un banco NO alcanzaría. Acá sí, porque:

1. **El QR no caduca automáticamente**: si alguien fuerza bruta un token, tiene que pasar `62^16` intentos contra un Sheet. Cada intento es una lectura del Sheet + renderizado de HTML → ~200ms por intento. Un atacante haciendo 5 intentos/segundo tardaría `3 × 10^20` años en promedio. 
2. **El `ADMIN_KEY` está separado**: el endpoint que da de alta vehículos (el único que podría generar QRs) está protegido por una clave separada de 32+ caracteres que NUNCA viaja por URL.
3. **El QR es físico**: pegado en el parabrisas. Si alguien quiere ver un historial, lo escanea. No necesita adivinar tokens.

Si en el futuro se necesita un CSPRNG (por ejemplo, para tokens con poder de revocación inmediata), se reemplaza `Math.random()` por `crypto.getRandomValues()` en el cliente o se llama a una API externa.

## Modelo de datos

Cinco hojas en un único Google Sheet — más dos hojas de catálogo curado que respaldan el dropdown de Marca/Modelo/Versión en `AdminPage.html`:

| Hoja                | Propósito                                                            |
|---------------------|----------------------------------------------------------------------|
| `Clientes`          | Datos del dueño (entidad débil)                                      |
| `Vehiculos`         | Entidad central: vehículo + FK lógica a cliente                      |
| `Servicios`         | Hechos: cada service es una fila                                     |
| `AccesosLog`        | Auditoría (sólo últimos 4 chars del token)                           |
| `Turnos`            | Agenda de servicios programados                                      |
| `Marcas_Modelos`    | Catalog (seed + user + migration entries) — backing dropdown para alta de vehículo |
| `Modelos_Versiones` | Curated versión list por (Marca, Modelo) — popula el tercer `<select>` |

### Catalog: aliases y normalización

El catálogo de Marcas/Modelos/Versión está respaldado por dos hojas (`Marcas_Modelos` y `Modelos_Versiones`) y se sirve a `AdminPage.html` a través de 3 selects en cascada en `#paso-vehiculo`. El cache vive en `CacheService.getScriptCache()` bajo una sola key `catalogoVehiculos_v1` con la forma `{ts, marcas, modelosPorMarca, versionesPorModelo}`. TTL de **6 horas (21600 s)** — bounded por el límite de 100 KB por entry de Apps Script; con el seed argentino (~80 marcas) el árbol estimado es 30-50 KB. En cada `registrarMarcaModelo()` exitoso se hace `cache.remove('catalogoVehiculos_v1')` (invalidación total — el upsert es raro, la simplicidad gana).

**Pipeline de normalización** (orden estricto): `trim` → collapse whitespace → Title Case → alias map sobre la primera palabra. La función pura `normalizarMarcaModelo(s)` vive duplicada en `Codigo.gs:2277` (canonical de backend, R2) y en `AdminPage.html` (pre-submit, A7) — están comentadas `Q2: duplicada` para mantenerlas sincronizadas.

**Aliases cubiertos**:

| Entrada  | Canónico       |
|----------|----------------|
| `vw`     | Volkswagen     |
| `chevro` | Chevrolet      |
| `chevr`  | Chevrolet      |
| `merced` | Mercedes-Benz  |
| `bmw`    | BMW            |
| `citroen`| Citroën        |
| `mb`     | Mercedes-Benz  |

**Auto-grow + rollback**: cada `altaVehiculo` y `altaVehiculoInterno` dispara `_autoGrowCatalogo_(marca, modelo, version)` que intenta `registrarMarcaModelo({origen: 'user'})` en `try/catch` independiente (R4). Una falla del catálogo NO bloquea el alta del vehículo — sólo loguea. La siembra inicial se hace vía menú **Catálogo → Cargar catálogo…** en el editor de Sheets (`cargarCatalogoMarcasModelos()`, ≥80 marcas + versiones curadas para Hilux/Etios/Ranger/Corolla/Focus/Cronos). La migración legacy aplica normalización + aliases con `Origen='migration'` vía menú **Catálogo → Migrar marcas…**. El rollback completo (eliminar hojas + invalidar cache) está en **Catálogo → Eliminar catálogo** (`eliminarHojasCatalogo()`).

### `Clientes` (entidad débil: solo para tener el nombre del dueño)

```
ID_Cliente | Nombre          | Telefono   | Email           | Fecha_Alta          | Consentimiento
C-0001     | Juan Pérez      | 1145678901 | juan@mail.com   | 2024-03-15 10:30:00 | SI
```

- `ID_Cliente`: generado por `nextId('Clientes', 'C')` → `C-0001`, `C-0002`, etc.
- `Consentimiento`: `SI` / `NO` — para cumplir con la Ley 25.326 de protección de datos personales.

### `Vehiculos` (la entidad central)

```
ID_Vehiculo | Token              | Fecha_Alta          | Activo | Patente | Marca  | Modelo | Anio | Combustible | ID_Cliente | URL_QR_Impresa
V-0001      | a7f9k2m8x4p1q6r3   | 2024-03-15 10:30:00 | TRUE   | AB123CD | Toyota | Etios  | 2020 | Nafta       | C-0001     | https://drive.google.com/...
```

- `Token`: 16 caracteres alfanuméricos aleatorios. Es la **única** pieza de información que va en el QR.
- `Activo`: cuando se setea a `FALSE`, el QR deja de funcionar. Sirve para dar de baja lógica.
- `Patente`: dato interno del taller, **NO** aparece en la URL ni en el HTML del cliente. Solo se muestra en la ficha del taller.
- `URL_QR_Impresa`: link al PNG del QR guardado en Drive, listo para imprimir.

### `Servicios` (los hechos: cada service es una fila)

```
ID_Servicio | Fecha       | ID_Vehiculo | Kilometraje | Descripcion              | Repuestos         | Proximo_Mantenimiento | Observaciones     | Fotos_IDs                                  | ID_Mecanico
S-0001      | 2024-03-15  | V-0001      | 50000       | Cambio de aceite y filtro| Filtro de aceite 1L| Cada 10.000 km        | Aceite Mobil 5W30| 1aBcD-2eFgH-3iJkL,4mNoP-5qRsT-6uVwX      | M-001
```

- `ID_Vehiculo`: FK lógica a `Vehiculos`. **No hay JOIN físico** — Apps Script hace la búsqueda leyendo ambas hojas (volúmenes pequeños).
- `Kilometraje`: lectura del odómetro al momento del service. Permite graficar evolución.
- `Proximo_Mantenimiento`: texto libre (ej. "Cada 10.000 km" o "2025-03-15"). Aparece como alerta amarilla arriba del historial.
- `Fotos_IDs`: lista de IDs de archivos de Drive, separados por coma. El HTML los divide y renderiza como thumbnails.

### `AccesosLog` (auditoría)

```
Timestamp            | Token_Suffix | Email_O_Anonimo
2024-03-15 14:22:01  | q6r3         | anonimo
2024-03-15 14:25:18  | q6r3         | anonimo
```

- Solo se loguean los últimos 4 caracteres del token (no el token entero, por seguridad).
- En el futuro se puede enriquecer con la IP del request (`e.parameter.remoteAddress` o headers HTTP) y un user-agent.
- Sirve para detectar accesos anómalos (ej. un token que recibe 1000 hits por día → algo raro).

### `Turnos` (agenda de servicios programados)

```
ID_Turno | Fecha_Hora          | Duracion_Minutos | ID_Vehiculo | Tipo_Servicio  | Descripcion                  | Estado      | ID_Mecanico | Fecha_Creacion       | Notas
T-0001   | 2026-06-30 14:30:00 | 60               | V-0001      | Cambio aceite  | Cliente pide revisar frenos  | pendiente   |             | 2026-06-28 10:00:00 | Traer llave repuesto
T-0002   | 2026-07-02 09:00:00 | 120              | V-0007      | Service completo | Service según manual       | confirmado  |             | 2026-06-25 11:00:00 |
```

- `ID_Turno`: generado por `nextId('Turnos', 'T')` → `T-0001`, `T-0002`, etc.
- `Fecha_Hora`: `Date` real en la celda (Google Sheets lo formatea como `yyyy-MM-dd HH:mm`). Se lee como string `yyyy-MM-dd HH:mm:ss` por `sheetAObjetos`.
- `Duracion_Minutos`: bloque estimado del turno en minutos (default 60, múltiplos de 15).
- `ID_Vehiculo`: FK lógica a `Vehiculos`. La función `listarTurnos(filtros)` enriquece cada turno con `Patente`, `Marca`, `Modelo`, `clienteNombre` y `Telefono_Cliente` para que el frontend no tenga que hacer N llamadas.
- `Tipo_Servicio`: enum cerrado (11 valores: Cambio de aceite, Service completo, Diagnostico, Freno, Embrague, Distribucion, Suspension, Tren delantero, Aire acondicionado, Bateria, Otro). El modal del admin usa un `<select>` para forzar valores válidos.
- `Estado`: enum cerrado de 5 valores:
  - `pendiente` (amarillo `#f59e0b`) — turno recién creado, esperando confirmación del cliente.
  - `confirmado` (azul `#3b82f6`) — el cliente confirmó asistencia.
  - `completado` (verde `#10b981`) — el service se hizo. Side effect: crea automáticamente una fila en `Servicios` (ver "Flujo Agenda → Servicios" abajo).
  - `cancelado` (gris `#6b7280`) — el turno no se va a hacer (cliente canceló, etc.).
  - `no_show` (rojo `#ef4444`) — el cliente no se presentó.
- `ID_Mecanico`: FK opcional a futuro (no implementado todavía, queda en blanco).
- `Fecha_Creacion`: timestamp de cuándo se agendó el turno. Sirve para auditoría ("¿cuándo se cargó este turno en la agenda?").
- `Notas`: texto libre, interno del taller (ej: "Traer llave de repuesto", "Cliente molesto, atender primero").

### Flujo Agenda → Servicios

El feature central de los turnos es que **completar un turno crea un servicio automáticamente**:

```
Usuario marca turno como "completado" en el modal
   ↓
Backend: completarTurnoInterno(idTurno)
   ↓
1. Actualiza Turnos.Estado = "completado" (fila del turno)
   ↓
2. Llama altaServicioInterno({
     idVehiculo: turno.ID_Vehiculo,
     fecha: turno.Fecha_Hora,
     kilometraje: 0,           ← placeholder, se completa cuando el cliente deja el auto
     descripcion: turno.Descripcion || turno.Tipo_Servicio,
     repuestos: '',             ← se completa post-service
     proximoMantenimiento: '',  ← se completa post-service
     observaciones: 'Generado automáticamente desde turno T-XXXX',
     fotosIds: '',
     idMecanico: turno.ID_Mecanico
   })
   ↓
3. Devuelve { ok: true, servicioId: "S-XXXX", message: ... }
```

**¿Por qué el side effect existe?**
Porque en el flujo real del taller:
1. Cliente reserva turno (se crea fila en Turnos).
2. Cliente llega el día del turno (se marca como completado).
3. Mecánico hace el service, sube fotos, completa kilometraje y repuestos (se actualiza el Servicio recién creado).

Sin el side effect, el mecánico tendría que cargar el servicio a mano después de cada turno, lo que duplica trabajo y abre la puerta a que se olvide.

**¿Por qué no usar `actualizarTurno` para esto?**
`actualizarTurno` es un update genérico. El side effect de creación de Servicio está atado específicamente a la acción "completado" (que es semánticamente distinta a "edité la fecha y de paso le puse completado"). Si en el futuro se setea completado vía el dropdown de estado del modal, NO se crea Servicio — solo se setea el estado. Para garantizar la creación hay que usar el botón explícito "Marcar completado".

**Decisiones de diseño relevantes:**
- **Kilometraje inicial = 0**: el auto todavía no está físicamente en el taller cuando se completa el turno (el turno es la reserva). Cuando el cliente deja el auto, el mecánico edita el Servicio y carga el km real.
- **Repuestos vacío, Próximo vacío**: lo mismo, son datos post-service.
- **Observaciones con el ID del turno**: para auditoría, queda registro de que el Servicio vino de un Turno.

## Modelo de seguridad

### Separación público / admin

El sistema tiene **dos endpoints** bien separados:

| Endpoint | Acceso | Qué hace |
|---|---|---|
| `doGet(?t=TOKEN)` | Público (sin auth) | Lee vehículo + servicios, renderiza HTML. **Solo lectura**. |
| `doPost(action=..., adminKey=...)` | Privado (requiere `ADMIN_KEY`) | Da de alta clientes, vehículos, servicios. Único endpoint de escritura. |

El `ADMIN_KEY`:
- Vive en `Script Properties` (nunca en el código fuente).
- Viaja en el body del POST, **nunca** en la URL (la URL queda en logs de proxies).
- Es de 32+ caracteres hex.
- Si se compromete, se rota: cambiar el valor en Script Properties → re-deploy → todos los Forms del taller se actualizan con el nuevo valor (guardado también en su Script Property).

### Por qué NO usar Google OAuth para el admin

- Los Forms son operados por el personal del taller, que no necesariamente tiene cuenta de Google (o tiene una personal y se olvida).
- El `ADMIN_KEY` se guarda **una vez** en el Apps Script del trigger del Form, y de ahí en más nadie lo tipea.
- Si el taller tiene 3 mecánicos, los 3 usan el mismo Form que dispara el mismo trigger con el mismo key. No hace falta OAuth individual.

### Por qué el QR es público (drive sharing `ANYONE_WITH_LINK`)

- El cliente que escanea el QR necesita ver las fotos SIN tener cuenta de Google.
- Drive con `ANYONE_WITH_LINK` + `VIEW` es el equivalente a "URL secreta" en Google. Si tenés el link, ves; si no, no existís.
- Las fotos NO son indexadas por Google (porque la URL nunca se publica), así que no aparecen en búsquedas.

### Defensa en profundidad

| Capa | Mitigación |
|---|---|
| Enumeración | Token aleatorio de 16 chars → espacio 62^16 |
| Fuga de patente | La patente NUNCA aparece en la URL ni en el HTML público |
| Inyección en HTML | Apps Script `<?= ?>` escapa HTML por defecto; los `<? if ?>` se evalúan server-side |
| Robar sesión | No hay sesión. Cada GET es stateless. El QR es el "ticket". |
| Subida maliciosa | El trigger del Form puede validar formato de archivos antes de aceptar |
| DoS | Google rate-limita Apps Script automáticamente. 30 GB/día de bandwidth en Web Apps. |

## Feature: Agenda de turnos

El taller necesita coordinar cuándo viene cada cliente. Antes de la Agenda, el mecánico anotaba los turnos en una planilla aparte o en el celular, sin relación directa con los vehículos del sistema. El feature agrega:

1. Una nueva hoja `Turnos` para registrar reservas.
2. Una sección "Agenda" en el panel admin con dos vistas: calendario mensual y lista de próximos 7 días.
3. Un modal de alta/edición con todos los datos del turno (vehículo, fecha/hora, duración, tipo, estado, notas).
4. Side effect automático: completar un turno crea el `Servicio` asociado.

### Endpoints del backend

| Acción API | Función | Propósito |
|---|---|---|
| `altaTurno` | `altaTurnoInterno(payload)` | Crea un turno nuevo. Valida vehículo + fechaHora + tipoServicio. |
| `listarTurnos` | `listarTurnos(filtros)` | Devuelve array de turnos enriquecidos con datos de vehículo y cliente. Acepta filtros opcionales `desde`, `hasta`, `estado`, `idVehiculo`. |
| `actualizarTurno` | `actualizarTurnoInterno(payload)` | Update parcial (solo los campos presentes en el payload). No genera side effects. |
| `cancelarTurno` | `cancelarTurnoInterno(id)` | Set atómico de `Estado = 'cancelado'`. |
| `completarTurno` | `completarTurnoInterno(id)` | Set `Estado = 'completado'` + crea fila en `Servicios`. |

### Vistas del frontend

**Calendario mensual** (`agendaView === 'calendario'`):
- Grilla de 42 celdas (6 filas × 7 columnas) con lunes como primer día.
- Cada celda muestra hasta 3 turnos como "chips" de color según estado + día del mes.
- `+N más` si hay más de 3 turnos en un día.
- Click en celda vacía abre el modal con esa fecha pre-cargada.
- Click en un chip abre el modal de edición de ese turno.
- Navegación: botones `‹` `›`, botón "Hoy" para volver al mes actual.
- Día actual destacado con borde rojo y número dentro de círculo rojo.
- Celdas de otros meses opacadas al 40%.
- Contador "X turnos este mes" arriba a la derecha.

**Lista próximos 7 días** (`agendaView === 'lista'`):
- 7 secciones (Hoy, Mañana, luego días de la semana) cada una con su fecha.
- Dentro de cada sección, tarjetas con: hora grande (Oswald), tipo de servicio, patente+modelo del vehículo, nombre del cliente + teléfono, badge de estado.
- Si una sección no tiene turnos muestra "Sin turnos" en gris.
- Click en una tarjeta abre el modal de edición.

**Modal de turno** (`#turnoModal`):
- Backdrop oscuro + blur, card centrada con scroll si excede viewport.
- Buscador de vehículo (mismo patrón que el buscador del servicio): typeahead con debounce de 250ms que llama a `buscarVehiculos`.
- Una vez seleccionado el vehículo, se oculta el buscador y aparece una card verde con el vehículo + nombre del cliente + botón "Cambiar".
- Campos: fecha, hora, duración (default 60, múltiplos de 15), estado (default `pendiente`), tipo (select cerrado de 11 valores), descripción, notas.
- Botones:
  - **Guardar** (siempre visible): valida y llama a `altaTurno` o `actualizarTurno` según corresponda.
  - **Marcar completado** (solo en edición, oculto si ya está completado/cancelado): confirm vía `confirm()` y llama a `completarTurno`. Toast incluye el ID del Servicio creado.
  - **Cancelar turno** (solo en edición, oculto si ya está cancelado): confirm vía `confirm()` y llama a `cancelarTurno`.
- Escape cierra el modal. Click en el backdrop también.
- Cuando se crea o edita un turno, se recarga la agenda con `loadAgenda()` que re-fetch todos los turnos y re-renderiza la vista activa.

### Demo y migración

**Función `cargarTurnosDemo()`**: genera 20 turnos distribuidos en los próximos 30 días usando los vehículos que ya estén cargados. Distribución:
- 4 turnos pasados en estado `completado` (servicios ya hechos).
- 1 turno pasado `cancelado` + 1 turno pasado `no_show` (para ver los colores).
- 5 turnos de hoy a +3 días en estado `confirmado` (reservas próximas).
- 9 turnos de +4 a +28 días en estado `pendiente` (reservas lejanas).

Total: 20 turnos. Tiene guard: si ya hay turnos, devuelve error pidiendo limpiar primero.

**Función `limpiarTurnosDemo()`**: borra todas las filas de la hoja `Turnos` (a diferencia de `limpiarDatosDemo()` que solo borra las huérfanas).

**Función `crearHojaTurnos()`**: idempotente. Crea la hoja `Turnos` con headers y formato si no existe; normaliza los headers si ya existe pero están mal. Pensada para correr una sola vez después del deploy del feature en sistemas existentes.

**`limpiarDatosDemo()` extendido**: ahora también borra turnos cuyo `ID_Vehiculo` quedó huérfano tras la limpieza de los vehículos demo. No borra turnos de vehículos reales.

### Consideraciones técnicas

- **Carga completa de turnos**: por simplicidad, `loadAgenda()` trae todos los turnos sin filtro de fecha. Para un taller con <1000 turnos esto es instantáneo. Si crece, se puede agregar filtro `desde/hasta` por mes visible.
- **Timezone**: Apps Script usa `Session.getScriptTimeZone()` para formatear las fechas al guardar. El frontend parsea con `new Date(str)` que usa la timezone del navegador. En la práctica (Argentina, GMT-3) no hay desfase perceptible.
- **No hay locks transaccionales**: `completarTurno` primero actualiza el turno y después crea el Servicio. Si la creación falla, el turno queda marcado como completado sin servicio asociado. Se documenta en el toast: el usuario puede crear el Servicio a mano desde la sección Servicio si pasa esto. Para el volumen esperado es aceptable.
- **Sin recurrencia**: cada turno es un evento único. Si el cliente viene todos los meses, hay que crear 12 turnos a mano. Se podría agregar recurrencia como feature futuro.

## Feature: Sistema de Presupuestos Digitales

El mecánico necesita poder armar un presupuesto (mano de obra + repuestos), enviarlo al cliente por WhatsApp/email, y que el cliente lo apruebe o rechace desde el celular. Si aprueba, se crea automáticamente un Servicio en el sistema. Es el flujo natural de cualquier taller: presupuesto → aprobación → ejecución.

### Por qué dos hojas (no JSON en una columna)

Una primera versión del feature guardaba los items como un JSON string en una sola hoja `Presupuestos`. Eso funcionaba para 1-2 items pero no escalaba:

- **Filtrar / buscar**: imposible buscar "qué presupuestos tienen filtro de aceite" sin parsear JSON en todas las filas.
- **Reportes**: "¿cuánto facturamos en filtros el último mes?" requería parsear JSON. Con dos hojas normalizadas es un `SUMIF` normal.
- **Edición**: para cambiar un item había que reescribir el JSON entero. Con la hoja `Presupuestos_Items`, se edita la fila directamente.
- **Volumen esperado**: 50-200 presupuestos/mes en un taller chico, con 3-10 items cada uno = 500-2000 filas por mes. Trivial para Sheets.

Decisión: **dos hojas normalizadas con FK lógica**, igual que `Turnos` (que también podría tener items pero por ahora es 1:1 con vehículo).

### Modelo de datos

#### `Presupuestos` (header)

```
ID_Presupuesto | Fecha              | Fecha_Vencimiento | ID_Vehiculo | Estado      | Subtotal | IVA    | Total   | Validez_Dias | Token_Publico       | Notas                       | ID_Servicio | Fecha_Aprobacion
P-0001         | 2026-06-28 10:00:00| 2026-07-13        | V-0001      | aprobado    | 25000    | 5250   | 30250   | 15           | a7f9k2m8x4p1q6r3   | Cliente frecuente           | S-0042      | 2026-06-29 14:30
```

- `ID_Presupuesto`: `P-0001`, generado por `nextId('Presupuestos', 'P')`.
- `Fecha`: timestamp de creación. Editable.
- `Fecha_Vencimiento`: `Fecha + Validez_Dias`, calculada al crear. Sirve para marcar como `vencido` si pasa.
- `ID_Vehiculo`: FK a `Vehiculos`. La lista de presupuestos enriquece con patente, marca, modelo, cliente y teléfono del vehículo.
- `Estado`: enum cerrado de 6 valores:
  - `borrador` (amarillo) — el mecánico lo está armando, todavía no se envió.
  - `enviado` (azul) — se marcó como enviado, el cliente ya tiene la URL.
  - `aprobado` (verde) — el cliente aprobó. Side effect: crea `Servicios`.
  - `rechazado` (rojo) — el cliente rechazó. El motivo se appendea a `Notas` con prefijo `[Rechazado: ...]`.
  - `vencido` (gris) — pasó `Fecha_Vencimiento` y sigue en `enviado`. **Estado calculado en lectura**, no se persiste (la celda queda con `enviado`, el cálculo se hace en `calcularEstadoPresupuesto()` y `listarPresupuestos` agrega el atributo `_EstadoCalculado`).
  - `completado` (emerald) — el servicio ya se hizo. Se setea manualmente desde el panel admin.
- `Subtotal` / `IVA` / `Total`: pre-calculados al crear (no se recalculan al editar). IVA = 21% (Argentina).
- `Validez_Dias`: cuántos días es válido. Default 15.
- `Token_Publico`: 16 chars alfanuméricos aleatorios, **diferente del token del vehículo**. Sirve para la URL pública del presupuesto. Se genera al crear el presupuesto con `generarTokenPresupuestoUnico()` y no cambia nunca.
- `Notas`: texto libre. Si el cliente rechaza, se appendea `[Rechazado: motivo]`.
- `ID_Servicio`: FK a `Servicios`. Se setea automáticamente al aprobar (side effect).
- `Fecha_Aprobacion`: timestamp de cuándo el cliente aprobó. Diferente de `Fecha` (creación).

#### `Presupuestos_Items` (line items, normalizada)

```
ID_Item | ID_Presupuesto | Descripcion                  | Cantidad | Precio_Unitario | Subtotal_Item | Tipo
PI-0001 | P-0001         | Filtro de aceite Mann W712/75| 1        | 5000            | 5000          | repuesto
PI-0002 | P-0001         | Mano de obra service         | 1        | 20000           | 20000         | mano_obra
```

- `ID_Item`: `PI-0001`, generado por `nextId('Presupuestos_Items', 'PI')`.
- `ID_Presupuesto`: FK al header.
- `Descripcion`: texto libre del item (ej: "Filtro de aceite Mann W712/75", "Mano de obra cambio de aceite").
- `Cantidad`: número. Puede ser decimal (1.5 hs de mano de obra, etc.).
- `Precio_Unitario`: precio por unidad.
- `Subtotal_Item`: `Cantidad * Precio_Unitario`, calculado al crear. Se persiste para no recalcular en cada lectura.
- `Tipo`: enum cerrado de 2 valores: `repuesto` o `mano_obra`. Sirve para mostrar icono/color distinto en la UI y para el texto del Servicio generado.

### Flujo end-to-end

```
1. Mecánico crea presupuesto desde el panel admin
   → altaPresupuestoInterno() genera P-0001, items PI-0001..N, Token_Publico
   → Estado = borrador

2. Mecánico revisa, agrega items, edita precios
   → sigue en borrador, no se envía todavía

3. Mecánico presiona "Enviar al cliente"
   → marcarPresupuestoEnviado() cambia estado a enviado
   → devuelve urlPublica (ya estaba generada al crear el presupuesto)

4. Mecánico copia URL o usa el botón "Enviar por WhatsApp"
   → la app genera un mensaje pre-formateado con todos los detalles + URL
   → wa.me/...?text=... se abre en nueva tab

5. Cliente abre la URL en su celular
   → PresupuestoPage.html se sirve desde el mismo Apps Script Web App
   → ve header con badge "enviado", items, totales
   → tiene botones "Aprobar" y "Rechazar"

6a. Cliente aprueba
   → google.script.run.aprobarPresupuesto({ token })
   → estado = aprobado
   → side effect: altaServicioInterno() crea S-0042 automáticamente
   → se guarda ID_Servicio en la fila del presupuesto
   → la página muestra confirmación "OK, servicio S-0042 generado"

6b. Cliente rechaza (con motivo opcional)
   → google.script.run.rechazarPresupuesto({ token, motivo })
   → estado = rechazado
   → se appendea "[Rechazado: motivo]" a Notas
   → la página muestra confirmación "X, gracias por avisarnos"

7. Mecánico ve el cambio en el panel admin (refresca o vuelve a la tab)
   → la card del presupuesto muestra el nuevo estado + ID_Servicio si fue aprobado

8. Cuando el service se ejecuta (via Turno completado o manual)
   → estado se setea a "completado" desde el panel admin
```

### Side effect: aprobar presupuesto crea Servicio

Mismo patrón que "completar Turno crea Servicio" del feature anterior:

```
aprobarPresupuesto(token):
  1. Buscar presupuesto por Token_Publico
  2. Validar estado === 'enviado' || 'borrador'
  3. Listar items del presupuesto
  4. Construir descripcion: "Presupuesto aprobado #P-XXXX · {items joined}"
  5. Construir repuestos: "{cant}x {desc}, {cant}x {desc}, ..."
  6. altaServicioInterno({
       idVehiculo, fecha=now, kilometraje=0,
       descripcion, repuestos,
       proximoMantenimiento='',
       observaciones='Generado automáticamente desde presupuesto P-XXXX',
       fotosIds='', idMecanico=''
     })
  7. Si servicio OK: actualizar presupuesto con estado=aprobado, ID_Servicio=servicio.id, Fecha_Aprobacion=now
  8. Devolver { ok, servicioId, message }
```

**Por qué el side effect existe**: el mecánico NO tiene que cargar el servicio a mano después de que el cliente aprueba. Es la misma lógica que `completarTurnoInterno()` — la aprobación es el evento que dispara la creación del trabajo a realizar.

**Decisiones de diseño**:
- `kilometraje: 0` — placeholder, se completa cuando el cliente deja el auto (igual que el caso del Turno).
- `repuestos` se llena con el formato `{cant}x {desc}` para que sea fácil de leer en el historial del cliente.
- `descripcion` lleva el ID del presupuesto (`#P-XXXX`) para auditoría y para que en el historial del vehículo sea identificable.
- `observaciones` deja registro: "Generado automáticamente desde presupuesto P-XXXX".

### Endpoints del backend

| Acción API | Función | Propósito |
|---|---|---|
| `altaPresupuesto` | `altaPresupuestoInterno(payload)` | Crea presupuesto + items. Valida vehículo, calcula totales, genera token. |
| `listarPresupuestos` | `listarPresupuestos(filtros)` | Array enriquecido (con vehículo y cliente). Acepta `estado`, `idVehiculo`, `desde`, `hasta`, `q`. Calcula estado `_EstadoCalculado` para detectar vencidos. |
| `obtenerPresupuesto` | `obtenerPresupuesto(id)` | Devuelve presupuesto + items + vehículo + cliente. |
| `actualizarPresupuesto` | `actualizarPresupuestoInterno(payload)` | Update parcial (estado, notas, idServicio, fechaAprobacion). |
| `marcarPresupuestoEnviado` | `marcarPresupuestoEnviado(id)` | Setea estado a `enviado`, devuelve `urlPublica`. |
| `aprobarPresupuesto` | `aprobarPresupuesto(token)` | Set estado `aprobado` + crea Servicio. **Usado desde la página pública**. |
| `rechazarPresupuesto` | `rechazarPresupuesto(token, motivo)` | Set estado `rechazado`, appendea motivo a notas. **Usado desde la página pública**. |
| `crearHojaPresupuestos` | `crearHojaPresupuestos()` | Idempotente. Crea las dos hojas si no existen. Pensada para deploy. |

### Endpoints públicos (sin sesión)

| URL | Función | Acceso |
|---|---|---|
| `?accion=presupuesto&token=XXXX` | `servePresupuestoPublico(token)` | Público. Renderiza `PresupuestoPage.html` con los datos del presupuesto. |

Las acciones de aprobar/rechazar se llaman desde la página servida por Apps Script vía `google.script.run`, que **no tiene problemas de CORS** (es la misma origin). El cliente NO necesita autenticarse — solo conocer el token, que es aleatorio y vive en la URL.

### Páginas del frontend

**AdminPage.html — sección `sec-presupuestos`**:
- Toolbar con filtro de estado + buscador libre + botón "+ Nuevo Presupuesto".
- Resumen con 5 stats: total, borradores, enviados, aprobados, monto activo (suma de totales excluyendo rechazados y vencidos).
- Cards con: badge de estado, ID + vehículo, cliente + fecha + vencimiento + cantidad de items + total destacado + acciones.
- Acciones por card: Detalle (abre modal), Enviar (solo borrador), Ver URL (enviado/vencido), Completado (solo aprobado).
- Modal "Nuevo Presupuesto": selector de vehículo (typeahead con debounce 300ms), validez en días, notas, items dinámicos con tipo/descripción/cantidad/precio/subtotal, totales en vivo (subtotal/IVA 21%/total).
- Modal "Detalle de Presupuesto": muestra todo en formato presentable + tabla de items + totales + mensaje pre-armado para WhatsApp + URL pública + acciones contextuales según estado.

**PresupuestoPage.html — página pública del cliente** (archivo separado):
- Mismo estilo dark mode Martinez que `ClientePage.html`.
- Header con logo, ID del presupuesto, fecha + vencimiento, badge de estado.
- Alert contextual según estado (vencido/aprobado/rechazado/completado).
- Cards con vehículo, cliente, items (tabla con descripción, cant, p.unit, subtotal, tipo), totales (subtotal/IVA/total destacado en rojo Oswald).
- Si el presupuesto está en `enviado` o `borrador`: botones "Aprobar presupuesto" (rojo grande) + "Rechazar" (gris). Rechazar abre textarea opcional para motivo.
- Confirmación inline (no redirige): "OK" verde para aprobado, "X" rojo para rechazado. La página no recarga — usa DOM updates.

### Demo y migración

**Función `cargarPresupuestosDemo()`**: genera 8 presupuestos distribuidos sobre los vehículos demo existentes:
- 3 en estado `borrador` (1 service 60k, 1 tren delantero, 1 frenos).
- 2 en estado `enviado` (1 bujias, 1 distribución completa).
- 2 en estado `aprobado` (cambio aceite + amortiguadores). Estos disparan `altaServicioInterno()` automáticamente, por lo que también se crean 2 Servicios nuevos con `kilometraje=0` y descripción que empieza con "Presupuesto aprobado #P-XXXX ·".
- 1 en estado `vencido` (batería). Como ya pasó la fecha de vencimiento, al leerlo se calcula `_EstadoCalculado=vencido`.

Cada presupuesto tiene 2-4 items mezclando `repuesto` y `mano_obra`. Total: ~24 items creados.

Guard: si ya hay presupuestos cargados, devuelve error pidiendo limpiar primero.

**Función `limpiarPresupuestosDemo()`**: borra todas las filas de `Presupuestos` y `Presupuestos_Items`. **No borra** los Servicios que se hayan creado por aprobaciones (esos quedan en la hoja `Servicios` como servicios normales con descripción que los identifica).

**Función `crearHojaPresupuestos()`**: idempotente. Crea las dos hojas con headers + formato (fechas como `yyyy-MM-dd`, números como `#,##0.00`, frozen first row). Pensada para correr una sola vez después del deploy.

**`limpiarDatosDemo()` extendido**: ahora también limpia presupuestos y sus items (bottom-up).

**Setup inicial extendido**: `setupInicial()` ahora crea las hojas `Presupuestos` y `Presupuestos_Items` si no existen (idempotente — se puede correr varias veces sin duplicar).

### Consideraciones técnicas

- **Token del presupuesto vs token del vehículo**: son dos cosas distintas. El token del vehículo (`?t=`) identifica al vehículo para ver el historial. El token del presupuesto (`?accion=presupuesto&token=`) identifica al presupuesto específico para aprobarlo/rechazarlo. Ambos son 16 chars alfanuméricos del mismo alphabet (`CHARSET`). Viven en hojas distintas (`Vehiculos.Token` vs `Presupuestos.Token_Publico`), así que no hay colisión.
- **Estado `vencido` calculado, no persistido**: para no tener que correr un cron job que actualice las filas, el cálculo se hace en cada lectura (`calcularEstadoPresupuesto()` y `listarPresupuestos` agrega `_EstadoCalculado`). El back-end siempre considera vencido si `estado === 'enviado' && Fecha_Vencimiento < now`. Trade-off: aceptamos el cálculo extra en lectura a cambio de no necesitar triggers.
- **Sin recurrencia**: un presupuesto es único. Si querés mandarle el mismo presupuesto al cliente todos los meses, hay que duplicarlo a mano. Está fuera de scope.
- **No hay versionado de items**: si el mecánico edita los items de un presupuesto aprobado, no se actualiza el Servicio asociado (el Servicio es un snapshot de lo aprobado). Es coherente con el patrón de "el Servicio refleja el trabajo aprobado".
- **PDF**: por ahora el "PDF" es la URL pública que el cliente puede abrir en el navegador y usar "Imprimir → Guardar como PDF". En una iteración futura se podría generar PDF server-side con Google Docs API o similar.

## Límites y consideraciones

| Límite | Valor | Impacto |
|---|---|---|
| Drive storage | 15 GB total | Suficiente para un taller chico-mediano por años. Cuando se llene, upgrade a Google One ($1.99/mes por 100 GB). |
| Apps Script runtime | 90 min/día | ~600 altas de vehículo por día. Suficiente para cualquier taller real. |
| Apps Script bandwidth | 30 GB/día en Web App | ~600.000 cargas de la página del cliente por día. |
| Sheet cells | 10M por hoja | ~250.000 servicios antes de quedarse sin espacio. |
| QuickChart | Sin rate limit conocido | No es un cuello de botella. |
| Apps Script execution time per call | 6 minutos | Suficiente para cualquier operación del sistema. |

## Decisiones de diseño que se descartaron

| Alternativa | Por qué NO |
|---|---|
| Firebase / Firestore | Costo a escala, requiere más setup, panel de admin extra. Sheets ya es conocido por el taller. |
| Supabase / Postgres | Requiere backend, deploy, mantenimiento. Overkill para este volumen. |
| Auth0 / OAuth | Fricción para el cliente que escanea (no quiere crear cuenta para ver su propio auto). |
| Patente en la URL | Enumerable, filtra PII. |
| QR firmado digitalmente | No aporta nada — el QR es solo un transportador de URL. La "firma" es el token aleatorio de 16 chars. |
| Hosting en Netlify / Vercel | Costo mensual + config de DNS + mantenimiento. Apps Script Web App es 1-click deploy. |
| Imágenes inline (base64) | Revienta el tamaño del HTML. Mejor thumbnails de Drive que cargan async. |
| PWA con service worker | Para este caso, la página es tan liviana que no vale la complejidad. |
