
# 📘 README – Bot Envío WhatsApp por VBA (Versión Actual + Mejoras)

## ✅ FUNCIONAMIENTO ACTUAL DEL BOT (`BOT_FUNCIONAL.txt`)

### 1. Lectura de datos desde hoja `Envio`
- Desde la fila 6 en adelante.
- **Columna A**: mensaje a enviar.
- **Columna C**: número de teléfono.
- **Columna D** (opcional): nombre del cliente.
- **Columna G**: status. Solo si dice `"ENVIAR"` se procede con el envío.
- **Columna H**: se marca como `"Enviado"` si fue exitoso.
- **Columna I**: errores o resultados adicionales.

### 2. Proceso paso a paso
1. Detecta ruta de Brave (64 o 32 bits).
2. Si `Status = "ENVIAR"`:
   - Toma el número y mensaje.
   - Agrega nombre si está disponible.
   - Codifica el mensaje para URL.
   - Abre WhatsApp Web en Brave con el mensaje y número prellenado.
   - Espera aleatoriamente 25–40 segundos para que cargue.
   - Envía el mensaje simulando `Enter` con `SendKeys "~"`.
   - Espera aleatoria 2–5 segundos y cierra Brave.
   - Marca como `"Enviado"` en hoja `Envio`.
3. Si hay error:
   - Se registra en la columna I.
   - Se cierra Brave y continúa.

---

## 🛠️ MEJORAS A IMPLEMENTAR

Queremos mantener todo el flujo actual intacto y simplemente agregar:

### 🔧 1. Hoja `Config`: tiempos personalizables
Permite configurar los intervalos de espera sin modificar el código. La hoja tendría esta estructura:

| Parámetro             | Min | Max |
|-----------------------|-----|-----|
| CargaWhatsApp         | 8   | 12  |
| EsperaAntesDeEnviar   | 8   | 12  |
| EsperaPostEnvio       | 5   | 7   |
| EsperaEntreClientes   | 12  | 18  |
| EnviosAntesDePausa    | 6   | 6   |
| DuracionPausaLarga    | 180 | 240 |

🔁 Uso en el código: el bot tomará un número aleatorio entre Min y Max de cada parámetro usando una función como `GetIntervalo("NombreDelParámetro")`.

### 📊 2. Hoja `Log`: registro por envío
Se creará un log automático que registre:

| Fecha y hora | Teléfono | Mensaje | TiempoTotal | Resultado | Error |
|--------------|----------|---------|-------------|-----------|-------|

Esto servirá para auditar qué fue enviado, cuándo, y en qué condiciones.

---

## ⚙️ FUNCIONAMIENTO PROPUESTO PASO A PASO (CON MEJORAS)

1. Leer hoja `Config` para obtener intervalos.
2. Recorrer fila por fila en `Envio` desde la fila 6.
3. Si `G = "ENVIAR"` (en vez de `"VENCIDO"`), continuar.
4. Leer número (columna C) y mensaje (columna A).
5. Agregar nombre (columna D), si aplica.
6. Codificar mensaje.
7. Construir URL con `https://web.whatsapp.com/send?...`.
8. Cerrar Brave si estaba abierto.
9. Esperar aleatoriamente (`CargaWhatsApp`).
10. Abrir Brave con URL.
11. Esperar (`EsperaAntesDeEnviar`).
12. Activar ventana y enviar `{ENTER}`.
13. Esperar (`EsperaPostEnvio`).
14. Registrar en hoja `Log`.
15. Esperar (`EsperaEntreClientes`).
16. Cada `EnviosAntesDePausa`, aplicar `DuracionPausaLarga`.
17. Marcar como "Enviado".
18. En caso de error, registrar en hoja `Log`.

---

## ✅ BENEFICIOS DE ESTAS MEJORAS

- 📊 Tiempos configurables: sin necesidad de editar el código.
- 🔒 Mayor naturalidad: evita bloqueos por comportamiento automatizado.
- 📝 Registro completo: para trazabilidad y control.
- 🔁 Más ordenado: separación de configuración, ejecución y log.

---

## 📊 ESTRUCTURA DE HOJAS EXCEL

### 📄 Hoja `Envio` (actual y mantenida)
| Columna | Contenido                     | Ejemplo              |
|---------|-------------------------------|----------------------|
| A       | Mensaje a enviar              | Hola {{nombre}}...   |
| B       | (vacía o auxiliar)            |                      |
| C       | Teléfono                      | 56998765432          |
| D       | Nombre del contacto (opcional)| Juan                 |
| E       | (no usada)                    |                      |
| F       | (no usada)                    |                      |
| G       | Estado del envío              | ENVIAR / Enviado     |
| H       | Resultado                     | Enviado / Error      |
| I       | Comentarios/Error             | Error de conexión... |

---

### ⚙️ Hoja `Config` (nueva para personalizar tiempos)

| Parámetro             | Min | Max |
|-----------------------|-----|-----|
| CargaWhatsApp         | 8   | 12  |
| EsperaAntesDeEnviar   | 8   | 12  |
| EsperaPostEnvio       | 5   | 7   |
| EsperaEntreClientes   | 12  | 18  |
| EnviosAntesDePausa    | 6   | 6   |
| DuracionPausaLarga    | 180 | 240 |

---

### 🧾 Hoja `Log` (estructura final para registrar actividad detallada)

| Columna              | Contenido                          | Ejemplo                          |
|----------------------|------------------------------------|----------------------------------|
| A – `Timestamp`       | Fecha y hora del envío             | `01/07/2025 15:23:10`            |
| B – `Fila origen`     | Número de fila procesada           | `6`                              |
| C – `Teléfono`        | Número de teléfono enviado         | `56998765432`                    |
| D – `Mensaje enviado` | Contenido del mensaje              | `Hola Juan, te escribo desde...` |
| E – `Tiempo carga (s)`| Tiempo de carga WhatsApp           | `10`                             |
| F – `Tiempo antes envío (s)` | Tiempo de espera antes de enviar  | `8`                              |
| G – `Tiempo post envío (s)`  | Tiempo de espera después del envío| `6`                              |
| H – `Resultado`       | Resultado del envío                | `Enviado` / `Error`              |
| I – `Nota/Error`      | Observaciones en caso de error     | `AppActivate falló`              |

