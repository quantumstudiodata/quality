# Quality CCI · Control de Calidad Interno de Parasitología

Aplicación web (Google Apps Script + Google Sheets) para el control de calidad diario del coproparasitoscópico. Cada laboratorio recibe al día una imagen positiva y una negativa. Un analista las identifica y el sistema califica automáticamente.

## Qué incluye esta versión (Fase 1)

- **Varios laboratorios.** Cada uno tiene su propio control diario, usuarios, banco de imágenes e historial.
- **Inicio de sesión con usuario y contraseña.** Las contraseñas se guardan cifradas. La contraseña inicial es temporal y se cambia de forma obligatoria al primer ingreso. Tras 5 intentos fallidos, la cuenta se bloquea 15 minutos.
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

## Códigos de usuario

| Laboratorio | Código | Supervisor | Analistas |
|---|---|---|---|
| Lab 01 Lister (origen) | LST | LST-SUP | LST-01, LST-02… |
| Lab 02 Azteca | AZT | AZT-SUP | AZT-01… |
| Lab 03 Swiss Lab | SWL | SWL-SUP | SWL-01… |
| Lab 04 Polab | PLB | PLB-SUP | PLB-01… |
| Lab 05 Jenner | JNR | JNR-SUP | JNR-01… |
| Lab 06 Moreira | MRA | MRA-SUP | MRA-01… |
| Lab 07 Liacsa | LCS | LCS-SUP | LCS-01… |
| Lab 08 Biomedica | BMD | BMD-SUP | BMD-01… |
| Lab 09 Exakta | EXK | EXK-SUP | EXK-01… |
| Lab 10 Promedic | PMD | PMD-SUP | PMD-01… |
| Lab 11 FamilyLabs | FML | FML-SUP | FML-01… |

Al inicio solo Lister está activo. Los demás laboratorios se activan desde la sección **Laboratorios** cuando decidan participar.

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
