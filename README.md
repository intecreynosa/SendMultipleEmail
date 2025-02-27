## **Envío de Correos desde un archivo Excel**

Este script en Python permite enviar correos electrónicos a varias direcciones contenidas en un archivo Excel. Personaliza el mensaje de correo de acuerdo con las columnas del archivo y permite definir si el correo debe enviarse como texto plano o como HTML.

### **Características**
- **Interfaz Gráfica**: Usa `tkinter` para permitir la configuración de los parámetros de correo, como la cuenta de correo, la contraseña, el servidor SMTP, el asunto y el cuerpo del mensaje.
- **Envío masivo de correos**: Permite enviar correos a múltiples destinatarios, cuyos correos están especificados en un archivo Excel.
- **Personalización del mensaje**: El cuerpo del correo puede personalizarse utilizando marcadores como `[C1]`, `[C2]`, etc., que se reemplazarán por los valores correspondientes de cada fila en el archivo Excel.
- **Configuración de HTML o texto plano**: El cuerpo del mensaje se puede enviar en formato HTML o en texto plano.
- **Selección de archivo Excel**: Puedes seleccionar el archivo Excel desde una ventana emergente.

### **Requisitos**
- Python 3.x
- Bibliotecas de Python: `tkinter`, `pandas`, `smtplib`, `email`
  - Puedes instalar las dependencias ejecutando:
    ```bash
    pip install pandas
    ```

### **Instrucciones de Uso**
1. **Abrir la aplicación**: Ejecuta el script para abrir la ventana de configuración.
2. **Configuración de Correo**:
   - Ingresa tu correo electrónico y contraseña en los campos correspondientes.
   - Asegúrate de usar un servidor SMTP válido (por ejemplo, `smtp.gmail.com` para Gmail).
3. **Archivo Excel**:
   - Selecciona el archivo Excel que contiene las direcciones de correo. Asegúrate de que la columna que contiene las direcciones de correo esté titulada "Email".
   - El archivo debe tener el formato `.xlsx`.
4. **Cuerpo del mensaje**:
   - Puedes personalizar el cuerpo del mensaje usando los valores de las columnas del archivo Excel, como `[C1]`, `[C2]`.
   - Si eliges que el correo se envíe en HTML, asegúrate de ingresar código HTML válido.
5. **Envío de correos**:
   - Haz clic en "Enviar Correos" para enviar los correos a todos los destinatarios del archivo Excel.

### **Ejemplo de archivo Excel**:
El archivo Excel debe tener al menos una columna titulada **"Email"** que contenga las direcciones de correo electrónico. Además, puedes agregar más columnas como **"C1"**, **"C2"**, etc., que se usarán para personalizar el mensaje.

**Ejemplo de archivo Excel**:

| Email            | C1    | C2    |
|------------------|-------|-------|
| usuario1@mail.com | Juan  | Pérez |
| usuario2@mail.com | Ana   | López |

### **Notas**:
- Asegúrate de no compartir tus credenciales (usuario y contraseña) con nadie.
- Si usas Gmail, es posible que necesites habilitar el acceso a aplicaciones menos seguras desde tu cuenta de Google.
- Si encuentras algún error, revisa los mensajes en la consola para detalles adicionales.
