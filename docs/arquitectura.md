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

Cuatro hojas en un único Google Sheet:

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
