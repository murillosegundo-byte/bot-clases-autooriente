# Bot de Seguimiento de Clases — CEA Auto Oriente

Sistema automatizado en Python para **CEA Auto Oriente** que captura los reportes fotográficos de clases prácticas enviados por los instructores en Telegram, almacena las evidencias en **Google Drive**, registra el seguimiento en **Google Sheets** y genera en memoria los reportes oficiales en PDF (**A.O.C.E.A-F-60, F-61 y F-62**) listos para imprimir y archivar en la carpeta del alumno.

---

## 1. Requisitos e Instalación

En una terminal de PowerShell dentro de `c:\Users\Usuario\Desktop\Bot\Practica_bot`:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

---

## 2. Configuración (`.env` y `credentials.json`)

1. Copia el archivo `.env.example` y renómbralo a `.env`:
   ```powershell
   Copy-Item .env.example .env
   ```
2. Coloca el archivo `credentials.json` (Cuenta de Servicio de Google Cloud con permisos de **Google Sheets API** y **Google Drive API**) en la raíz del proyecto.
3. Comparte tu documento de Google Sheets y la carpeta raíz de Google Drive con el correo de la cuenta de servicio (`client_email` dentro de `credentials.json`) como **Editor**.
4. (Opcional) Coloca el logotipo de la academia en `assets/logo_autooriente.png` para que se incluya automáticamente en el encabezado de los PDFs.

---

## 3. Estructura de Google Sheets

Al ejecutar el comando `/inicializar` en Telegram, el bot verifica o crea automáticamente las **3 hojas clave**:

1. **`Alumnos`**: `ID_Matricula` | `Documento` | `Nombre_Completo` | `Categoria` (`A2`, `B1`, `C1`, `Recat_C1`) | `Estado`
2. **`Instructores`**: `Telegram_User_ID` | `Nombre_Instructor` | `Categoria_Habilitada`
3. **`Registro_Clases`**: `Timestamp` | `Fecha_Reporte` | `Hora_Reporte` | `Instructor` | `ID_Alumno` | `Nombre_Alumno` | `Categoria` | `Clases_Reportadas` | `Tema_Registrado` | `URL_Foto_Drive`

---

## 4. Ejecución del Bot

```powershell
python main.py
```

### Comandos en Telegram
- **Foto + comentario con la palabra `Clase`**: Registra automáticamente la clase, sube la foto a Google Drive en la carpeta del alumno y responde con `✅`.
- `/reporte [Documento o Nombre del Alumno]`: Genera y envía al chat el PDF oficial (**F-60** para A2, **F-61** para B1, **F-62** para C1 / Recategorización C1).
- `/inicializar`: Crea las pestañas y encabezados en Google Sheets.
- `/ayuda`: Muestra la guía rápida en el chat.
