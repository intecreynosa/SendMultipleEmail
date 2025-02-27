import tkinter as tk
from tkinter import messagebox, filedialog
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

def enviar_correo(usuario, contrasena, servidor_smtp, asunto, cuerpo, html, archivo_excel):
    try:
        # Leer el archivo Excel
        df = pd.read_excel(archivo_excel)
        
        # Conectar al servidor SMTP de Gmail
        server = smtplib.SMTP(servidor_smtp, 587)
        server.starttls()
        server.login(usuario, contrasena)
        
        # Enviar el correo a cada fila del archivo Excel
        for index, row in df.iterrows():
            print(row)
            # Personalización del correo con la información de Excel
            correo_destino = row['Email']
            mensaje_personalizado = cuerpo
            
            # Reemplazar las variables como [C1], [C2] por los valores correspondientes
            for col in df.columns:
                mensaje_personalizado = mensaje_personalizado.replace(f"[{col}]", str(row[col]))

            # Crear el mensaje
            msg = MIMEMultipart()
            msg['From'] = usuario
            msg['To'] = correo_destino
            msg['Subject'] = asunto
            
            if html:
                # Si el correo es HTML
                msg.attach(MIMEText(mensaje_personalizado, 'html'))
            else:
                # Si el correo es en texto plano
                msg.attach(MIMEText(mensaje_personalizado, 'plain'))
            
            # Enviar el correo
            server.sendmail(usuario, correo_destino, msg.as_string())
        
        # Cerrar la conexión con el servidor SMTP
        server.quit()
        messagebox.showinfo("Éxito", "Correos enviados exitosamente.")
    
    except Exception as e:
        messagebox.showerror("Error", f"Hubo un error: {str(e)}")


def seleccionar_archivo():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo:
        archivo_excel_var.set(archivo)


# Crear la ventana principal
root = tk.Tk()
root.title("Envio de Correos")

# Variables
usuario_var = tk.StringVar()
contrasena_var = tk.StringVar()
servidor_smtp_var = tk.StringVar(value="smtp.gmail.com")
asunto_var = tk.StringVar()
cuerpo_var = tk.StringVar()
html_var = tk.BooleanVar()
archivo_excel_var = tk.StringVar()

# Crear los componentes de la interfaz gráfica
tk.Label(root, text="Correo (usuario):").grid(row=0, column=0)
tk.Entry(root, textvariable=usuario_var, width=40).grid(row=0, column=1)

tk.Label(root, text="Contraseña:").grid(row=1, column=0)
tk.Entry(root, textvariable=contrasena_var, show="*", width=40).grid(row=1, column=1)

tk.Label(root, text="Servidor SMTP:").grid(row=2, column=0)
tk.Entry(root, textvariable=servidor_smtp_var, width=40).grid(row=2, column=1)

tk.Label(root, text="Asunto:").grid(row=3, column=0)
tk.Entry(root, textvariable=asunto_var, width=40).grid(row=3, column=1)

tk.Label(root, text="Cuerpo del mensaje:").grid(row=4, column=0)
# Usamos un widget Text para permitir varias líneas
cuerpo_text = tk.Text(root, height=10, width=40)  # Ajusta el tamaño según lo necesites
cuerpo_text.grid(row=4, column=1)
cuerpo_text.insert(tk.END, "Hola [C1] [C2], ¿cómo estás?")  # Un valor inicial de ejemplo

tk.Checkbutton(root, text="HTML", variable=html_var).grid(row=5, column=0, columnspan=2)

tk.Label(root, text="Archivo Excel:").grid(row=6, column=0)
tk.Entry(root, textvariable=archivo_excel_var, width=40).grid(row=6, column=1)
tk.Button(root, text="Seleccionar", command=seleccionar_archivo).grid(row=6, column=2)

# Botón para enviar los correos
tk.Button(root, text="Enviar Correos", command=lambda: enviar_correo(
    usuario_var.get(),
    contrasena_var.get(),
    servidor_smtp_var.get(),
    asunto_var.get(),
    cuerpo_text.get("1.0", tk.END),  # Obtener todo el texto ingresado en el widget Text
    html_var.get(),
    archivo_excel_var.get()
)).grid(row=7, column=0, columnspan=3)

# Iniciar la ventana
root.mainloop()
