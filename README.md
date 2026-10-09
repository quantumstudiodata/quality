# Quality CCI · Control de Calidad Interno de Parasitología

Aplicación web (Google Apps Script + Google Sheets) para el control de calidad diario del coproparasitoscópico. Cada laboratorio recibe al día una imagen positiva y una negativa. Un analista las identifica y el sistema califica automáticamente.

## Qué incluye esta versión (Fase 1)

- **Varios laboratorios.** Cada uno tiene su propio control diario, usuarios, banco de imágenes e historial.
- **Inicio de sesión con usuario y contraseña.** Las contraseñas se guardan cifradas. La contraseña inicial es temporal y se cambia de forma obligatoria al primer ingreso. Tras 5 intentos fallidos, la cuenta se bloquea 15 minutos.
- **Ventana del control al entrar.** Al analista se le abre el control del día en una ventana que cubre toda la página; responde y envía ahí mismo. Puede minimizarla y retomarla desde "Estado de hoy" en el menú lateral.
- **Dos roles:**
  - **Analista:** responde el control diario y consulta el historial.
  - **Supervisor:** ve el control de hoy con las respuestas correctas, los indicadores y el historial; administra el banco de imágenes y las cuentas de sus analistas.
  - El supervisor de la marca de origen (`LST-SUP`) además administra los laboratorios y da de alta a sus supervisores.
- **Banco de imágenes en dos niveles:**
  - **Banco general:** las imágenes que sube Lister (marca de origen). Se distribuyen a todos los laboratorios.
  - **Banco propio:** las imágenes que sube cada laboratorio. Solo las usa ese laboratorio.
- **Rotación sin consumir imágenes.** Cada día se elige la imagen que el laboratorio no ha visto o la que vio hace más tiempo. Para no repetir en 90 días se recomiendan al menos 90 positivas y 90 negativas; el sistema avisa si hay menos.
- **Ciclo de vida.** A los **6 meses** de subida, la imagen sale de la rotación y se archiva. A los **9 meses** se borra definitivamente.
- **Un control por laboratorio y día**, con hasta 3 intentos compartidos. Queda registrado quién respondió y en qué intento.
- **Indicadores de los últimos 30 días:** porcentaje aprobado al primer intento, aprobados con reintento, no aprobados y días sin respuesta.
- **Seguridad:** el servidor valida el rol en cada acción y nunca envía al analista la respuesta correcta ni cuál imagen es la positiva.
- **Fechas por zona horaria del laboratorio**, protección contra accesos simultáneos y memoria temporal en el servidor para cargar más rápido.

## Usuarios

- Cada cuenta se da de alta con **nombre completo, iniciales, número de analista, puesto y laboratorio**.
- El usuario para iniciar sesión es **iniciales + número de analista**, por ejemplo `MPWN1` o `EMAR2`. Debe ser único en todos los laboratorios.
- Los supervisores no tienen número de analista: su usuario es **iniciales + `-SUP`**, por ejemplo `JLR-SUP`. El sufijo no se puede modificar.
- Los supervisores pueden editar sus propios datos y los de sus analistas. Solo la administración cambia el laboratorio de una cuenta.
- **Eliminar** una cuenta impide que vuelva a entrar y la quita de la lista; sus respuestas pasadas se conservan en el historial.
- Al crearla se genera una **contraseña temporal de 4 dígitos**. En su primer ingreso la persona crea la suya, también de **4 caracteres**, y puede cambiarla después desde **Mi perfil**.
- En **Mi perfil** cada quien ve su nombre, puesto, usuario y laboratorio, sin poder editarlos.
- La cuenta inicial de administración es `LST-SUP`.

Los laboratorios se muestran como `LST-01 Lister`, `AZT-02 Azteca`, etc. Al inicio solo Lister está activo; los demás se activan desde **Laboratorios** cuando decidan participar.

## Velocidad

Al iniciar sesión, la aplicación descarga en **una sola llamada** todo lo que necesita (control de hoy, indicadores, historial, atlas y, para el supervisor, banco y usuarios). Cambiar de pestaña es inmediato porque no vuelve a consultar al servidor. Después de cada acción, los datos se actualizan en segundo plano. Al recargar la página se muestra al instante lo último que se vio mientras se actualiza.

## Instalación

1. Crea una **Google Sheet nueva** (recomendado) o abre la que usas hoy. Los nombres de las hojas nuevas no chocan con los de la versión anterior.
2. Abre **Extensiones → Apps Script**.
3. Reemplaza el contenido de `Código.gs` con el de `Code.gs`. Crea un archivo HTML llamado `Index` y pega el contenido de `Index.html`.
4. Opcional: en **Configuración del proyecto**, activa “Mostrar el archivo de manifiesto appsscript.json” y pega el contenido de `appsscript.json`.
5. En el editor, elige la función **`configurarInicial`** y presiona **Ejecutar**. Acepta los permisos que pida Google.
   - En el **Registro de ejecución** aparecerá la contraseña temporal de `LST-SUP`. Guárdala.
   - Esta función crea las hojas, los 11 laboratorios, la cuenta `LST-SUP` y el disparador diario de mantenimiento (3:00 a.m.).
6. **Migrar tus datos actuales** (opcional, una sola vez):
   - Si tus datos están en **otra** hoja de cálculo, pega su ID en `ID_HOJA_ANTERIOR`, dentro de `Code.gs`. El ID es el texto entre `/d/` y `/edit` en la URL.
   - Ejecuta **`migrarDatosAnteriores`**. Tus imágenes (también las ya usadas, que vuelven a rotar), el historial de imágenes y las respuestas pasan a Lister. Las imágenes migradas cuentan sus 6 meses desde el día de la migración.
7. **Implementar → Nueva implementación → Aplicación web:**
   - Ejecutar como: **Yo**.
   - Quién tiene acceso: **Cualquier usuario**. El inicio de sesión de la propia app protege el acceso; no hace falta cuenta de Google.
8. Abre la URL, entra con `LST-SUP` y la contraseña temporal, y crea tu contraseña personal.

## Uso diario

- **Supervisor de Lister:** sube imágenes al banco general, da de alta laboratorios y a sus supervisores (sección **Usuarios**, rol Supervisor).
- **Supervisor de cada laboratorio:** da de alta a sus analistas, puede subir imágenes propias y revisa el control de hoy, el historial y los indicadores.
- **Analista:** entra, responde las dos imágenes del día (hasta 3 intentos) y consulta su historial.
- **¿Olvidaste tu contraseña?** El supervisor la restablece desde **Usuarios** y se genera una nueva temporal.

## Notas

- La sesión dura 6 horas. Al desactivar a un usuario, el cambio se aplica a más tardar al terminar su sesión actual.
- Las imágenes se publican en Drive como “cualquier persona con el enlace” para que carguen rápido. Los archivos se nombran con un código, no con el nombre del parásito.
- Si se cambia `Code.gs`, hay que crear una **nueva versión** de la implementación (Implementar → Administrar implementaciones → Editar → Nueva versión).
