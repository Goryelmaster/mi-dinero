# Conectar Mi dinero a tu Google (una sola vez, ~10 min)

Hazlo en la computadora. El paso 9 es en el iPhone.

## A. Pegar el script
1. Entra a **script.google.com** con tu cuenta de Google → **Nuevo proyecto**.
2. Arriba a la izquierda, cambia el nombre "Proyecto sin título" por **API Santino**.
3. Borra todo lo que hay en `Código.gs` y pega todo el archivo [`Codigo.gs`](Codigo.gs). Toca 💾 Guardar.
4. ⚙️ **Configuración del proyecto** (barra izquierda) → Zona horaria: **(GMT-07:00) Mazatlán**.

## B. Darle permisos y sacar tu clave
5. Vuelve a **Editor** (`< >`). Arriba, en la lista de funciones, elige **instalar** → toca **▶ Ejecutar**.
6. Te pide permisos: **Revisar permisos** → elige tu cuenta → sale "Google no verificó esta app" → **Configuración avanzada** → **Ir a API Santino (no seguro)** → **Permitir**.
   (Es normal: la app la hiciste tú y solo tú la usas. Pide Sheets, Drive y Gmail para guardar tus datos y leer recibos.)
7. Abajo, en **Registro de ejecución**, aparece: `Tu clave secreta es: …`. **Cópiala** (por ejemplo, en Notas, bloqueada). No la compartas.

## C. Publicarlo como API
8. Arriba a la derecha: **Implementar → Nueva implementación** → ⚙️ tipo **Aplicación web**:
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
   - **Implementar** → copia la **URL de la aplicación web** (termina en `/exec`).
   La URL sola no sirve de nada sin tu clave: sin ella no se puede leer ni guardar nada.

## D. Ponerlo en el iPhone
9. Abre **Mi dinero** desde tu pantalla de inicio. Ciérrala y ábrela otra vez (así baja la versión nueva).
10. Toca el engranaje ⚙️ → baja a **Nube (Google)**:
    - Pega la **URL** en "Dirección del script".
    - Pega la **clave** en "Clave secreta".
11. Arriba debe decir **✓ En la nube**. Si dice "⏳ Pendiente (error)", abajo de "Subir ahora" sale el motivo.

## E. Verificar que funciona
12. En tu Google Drive, abre la carpeta **Datos apps**: debe estar **mi-dinero-respaldo.json**, la hoja **Mi dinero - datos** y la carpeta **respaldos**.
13. Prueba sin internet: pon **modo avión** → anota un gasto → arriba dice "Sin internet · ⏳ 1 pendiente" → quita el modo avión → en unos segundos dice **✓ En la nube** y el gasto aparece en la pestaña `movimientos` de la hoja.
14. Gmail: en Sobres aparece **Revisar** con cobros que encontró en tu correo. Escribe cuánto te cobran en pesos → **Aceptar**, o **Ignorar**. Se revisa solo cada día a las 7 am.

## Actualizar el código
Cuando haya una versión nueva de `Codigo.gs` (por ejemplo, para el botón **Buscar recibos ahora**):
1. Abre tu proyecto **API Santino** en script.google.com, borra todo lo de `Código.gs` y pega el [`Codigo.gs`](Codigo.gs) nuevo → 💾 Guardar.
2. **Implementar → Administrar implementaciones** → ✏️ (lápiz) → Versión: **Nueva versión** → **Implementar**. La URL y la clave no cambian.
3. Si la app dice "Actualiza tu Apps Script", es que falta este paso.

## Para "Mi día" (después)
Usa la misma URL y la misma clave. La app manda `"app": "mi-dia"` y el script crea solo **Mi día - datos** y **mi-dia-respaldo.json** en la misma carpeta.

## Mi día → Google Calendar (automático)
1. Pega el `Codigo.gs` nuevo en tu proyecto **API Santino** → 💾 Guardar.
2. Arriba elige la función **sincronizarCalendario** → **▶ Ejecutar** → acepta el permiso nuevo de **Calendar** (mismos pasos del punto 6).
3. **Implementar → Administrar implementaciones → ✏️ → Versión: Nueva versión → Implementar.** La URL no cambia.
4. En Google Calendar aparece el calendario **Mi día** con tus empleados y tu rutina fija (aviso 5 min antes). Los 6 que se pusieron a mano en tu calendario principal se borran solos la primera vez.
5. Desde ahora, cuando cambias un horario en Mi día, el calendario se actualiza solo la próxima vez que la app sube a la nube.
