import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import locale
import re
import pandas as pd
import smtplib
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime
from cryptography.fernet import Fernet


smtp_server = "smtp.gmail.com"
smtp_port = 587
ciudades_pendientes={}
errores=[]
entregas_realizadas=0
entregas_procesadas=0

valor_total=0
CREDENCIALES_FILE = "credenciales.enc"
KEY_FILE = "clave.key"

"""Genera una clave si no existe y la guarda en un archivo."""
def generar_o_cargar_clave():
    if not os.path.exists(KEY_FILE):
        clave = Fernet.generate_key()
        with open(KEY_FILE, "wb") as key_file:
            key_file.write(clave)
    else:
        with open(KEY_FILE, "rb") as key_file:
            clave = key_file.read()
    return clave

CLAVE = generar_o_cargar_clave()
CIFRADOR = Fernet(CLAVE)

"""Cifra los datos con Fernet."""
def cifrar_datos(datos):
    return CIFRADOR.encrypt(datos.encode())


"""Descifra los datos con Fernet."""
def descifrar_datos(datos_cifrados):
    return CIFRADOR.decrypt(datos_cifrados).decode()


"""Guarda las credenciales cifradas si el checkbox está activado."""
def guardar_credenciales():
    if recordar_var.get():
        datos = f"{entry_email.get()}\n{entry_password.get()}"
        datos_cifrados = cifrar_datos(datos)
        with open(CREDENCIALES_FILE, "wb") as f:
            f.write(datos_cifrados)
    else:
        if os.path.exists(CREDENCIALES_FILE):
            os.remove(CREDENCIALES_FILE)
    

"""Carga y descifra credenciales si existen."""
def cargar_credenciales():
    if os.path.exists(CREDENCIALES_FILE):
        with open(CREDENCIALES_FILE, "rb") as f:
            datos_cifrados = f.read()
            try:
                datos = descifrar_datos(datos_cifrados).split("\n")
                entry_email.insert(0, datos[0])
                entry_password.insert(0, datos[1])
                recordar_var.set(1)
            except Exception:
                print("Error al descifrar los datos.")

"""Enviamos correos a los distintos clientes en base al estado de su pedido """
def EnviarAviso(Email, server, row):
    global valor_total, entregas_realizadas,entregas_procesadas

    Estado = row["Estado_Entrega"]
    Cliente = row["Cliente"]
    Id_Entrega = row["ID_Entrega"]
    Email_cliente = row["Correo_Cliente"]

    msg = MIMEMultipart()
    msg["From"] = Email
    msg["To"] = Email_cliente
    
    entregas_procesadas+=1
    if Estado == "Pendiente":
        Ciudad = row["Ciudad"]
        ciudades_pendientes[Ciudad] = ciudades_pendientes.get(Ciudad, 0) + 1
        asunto = "Tu pedido está en camino 🚚"
        cuerpo = f"Hola {Cliente},\nTu pedido con ID {Id_Entrega} está en camino y será entregado pronto."
    elif Estado == "Entregado":
        entregas_realizadas += 1
        valor_total+=row["Valor"]
        asunto = "Tu pedido ha sido entregado 🎉"
        cuerpo = f"Hola {Cliente},\nTu pedido con ID {Id_Entrega} ha sido entregado con éxito.\nGracias por confiar en nosotros."
    else:
        entregas_procesadas-=1
        errores.append(f"Estado inválido en la entrega {Id_Entrega}")
        return 

    msg["Subject"] = asunto
    msg.attach(MIMEText(cuerpo, "plain"))

    try:
        server.sendmail(Email, Email_cliente, msg.as_string())
    except Exception as e:
        
        errores.append(f"Error al enviar correo a {Email_cliente}: {e}")

locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

"""Formateamos la fecha segun los posibles tipos al formato deseado"""
def format_date(fecha):
    for formato in ['%Y/%m/%d', '%d-%m-%Y', '%B %d, %Y', '%Y-%m-%d %H:%M:%S']:
        try:
            fecha_formateada = datetime.strptime(fecha, formato).date()
            return fecha_formateada.strftime('%Y-%m-%d')
        except ValueError:
            continue
    return fecha
 
"""eliminamos los espaciados indeseados del nombre del cliente"""   
def format_client(cliente):
    cliente=re.sub(r"\s+", " ", cliente).strip()
    return cliente

"""formateamos el valor para que sea con . decimal y 2 de estos"""   
def format_valor(valor):
    if isinstance(valor, str):
        valor = valor.replace('.', '').replace(',', '.')
        return float(valor) 
    elif isinstance(valor, (int, float)):
        return float(valor) 
    else:
        raise ValueError(f"Tipo de dato no soportado: {type(valor)}")

"""aplicamos los distintos metodos al dataframe para que cumpla con el formato deseado"""   
def formatear_entregas(archivo):
    df = pd.read_excel(archivo)
    print("Columnas detectadas:", df.columns)
    
    df["Fecha_Pedido"] = df["Fecha_Pedido"].astype(str).apply(format_date)
    df["Cliente"] = df["Cliente"].apply(format_client) 
    df["Valor"] = df["Valor"].apply(format_valor) 
    df = df[~df["Estado_Entrega"].str.contains("Devuelto", case=False, na=False)]
    return df

"""se usa para cambiar el mensaje de estado"""   
def actualizar_estado(mensaje, color="black"):
    """Muestra el mensaje en la etiqueta de estado con color."""
    estado_label.config(text=mensaje, fg=color)
    root.update_idletasks()  # Ac     

"""Se gestiona el Archivo excel"""   
def seleccionar_entregas():
    archivo = filedialog.askopenfilename(title="Seleccionar archivo")
    if archivo:
        entry_archivo.delete(0, tk.END)
        entry_archivo.insert(0, archivo)
        
"""se procesa el archivo excel y itera para mandar los correos"""   
def procesar_entregas():
    email_sender = entry_email.get()
    email_password = entry_password.get()
    archivo = entry_archivo.get()
    
    if not email_sender or not archivo or not email_password:
        messagebox.showwarning("Advertencia", "Ingrese un correo, contraseña y archivo svalidos")
        actualizar_estado("⚠️ Ingrese datos válidos", "red")

        return
    
    df = formatear_entregas(archivo)
    df.to_excel("./entregas_procesadas.xlsx", index=False)

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()  # Cifra la conexión
    server.login(email_sender, email_password)
    
    try:
        actualizar_estado(" Conectando...", "blue")
        server = smtplib.SMTP("smtp.gmail.com", 587)  # Servidor SMTP de Gmail
        server.starttls() 
        server.login(email_sender, email_password)  # Inicia sesión
        print("✅ Conexión SMTP exitosa.")
        actualizar_estado("✅ Conectado. Enviando Correos...", "green")

        for _, row in df.iterrows():
            EnviarAviso(email_sender, server, row)
            
        server.quit()
    except smtplib.SMTPAuthenticationError as e:
        print(f"❌ Error de autenticación: {e}")
    except Exception as e:
        print(f"❌ Otro error: {e}")

    
    except Exception as e:
        errores.append("Error No se pudo conectar al servidor SMTP: {e}")
        

    descargar_reporte()
    
"""se descarga el reporte resumen con su informacion"""   
def descargar_reporte():
    ruta = "./reporte_resumen.txt"
    
    try:
        with open(ruta, "w") as file:
            file.write(f"Hubo un total de {entregas_procesadas} entregas procesadas.\n\n")
            
            if ciudades_pendientes:
                max_pendientes = max(ciudades_pendientes.values())  # Encuentra el número máximo de entregas pendientes
                ciudades_maximas = [ciudad for ciudad, cantidad in ciudades_pendientes.items() if cantidad == max_pendientes]

                if len(ciudades_maximas) == 1:
                    file.write(f"{ciudades_maximas[0]} es la ciudad con más entregas pendientes ({max_pendientes} entregas pendientes).\n\n")
                else:
                    file.write(f"Las ciudades con más entregas pendientes ({max_pendientes} entregas pendientes) son: {', '.join(ciudades_maximas)}.\n\n")
            
            file.write(f"Se realizaron un total de {entregas_realizadas} entregas exitosas.\n\n")

            file.write(f"Monto total de entregas realizadas. : {valor_total:.2f}\n\n")

            if errores:
                file.write("Errores encontrados:\n")
                for i in errores:
                    file.write(f"- {i}\n")
        
        print("Reporte guardado correctamente.")
        actualizar_estado("✅ Reporte generado con éxito.", "green")

    except Exception as e:
        actualizar_estado(f"❌ Error: {str(e)}", "red")
        print(f"Error al guardar el reporte: {e}")

        
"""DISEÑO"""        
root = tk.Tk()
root.title("Procesador de Datos")
root.geometry("450x400")
root.configure(bg="#f4f4f4")  # Fondo gris claro

# Estilos
label_style = {"font": ("Arial", 12, "bold"), "bg": "#f4f4f4"}
entry_style = {"font": ("Arial", 12), "width": 30, "bd": 2, "relief": "solid"}
button_style = {"font": ("Arial", 12, "bold"), "fg": "white", "bg": "#4CAF50", "bd": 2, "relief": "raised"}

# Espaciado
tk.Label(root, text="Ingrese un correo:", **label_style).pack(pady=5)
entry_email = tk.Entry(root, **entry_style)
entry_email.pack(pady=5)



tk.Label(root, text="Ingrese la contraseña:", **label_style).pack(pady=5)
entry_password = tk.Entry(root, show="*", **entry_style)  # Ocultar contraseña
entry_password.pack(pady=5)

recordar_var = tk.IntVar()
chk_recordar = tk.Checkbutton(root, text="Recordar credenciales", variable=recordar_var, font=("Arial", 10), bg="#f4f4f4", command=guardar_credenciales)
chk_recordar.pack(pady=5)

tk.Button(root, text="Seleccionar Archivo", command=seleccionar_entregas, **button_style).pack(pady=10)
entry_archivo = tk.Entry(root, **entry_style)
entry_archivo.pack(pady=5)

tk.Button(root, text="Procesar", command=procesar_entregas, bg="#008CBA", fg="white", font=("Arial", 12, "bold"), bd=2, relief="raised").pack(pady=20)

estado_label = tk.Label(root, text="Esperando datos...", font=("Arial", 12, "bold"), bg="#f4f4f4", fg="black")
estado_label.pack(pady=10)


cargar_credenciales()

root.mainloop()