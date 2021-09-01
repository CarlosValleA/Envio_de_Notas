from generar_informe import GenerarInforme
from enviar_por_mail import Mail


archivo = "notas.xlsx"
hoja = "T1_mail"
asunto = "Notas Tarea 1"  # Asunto del correo


# Funcion para generar los reportes. Si se quiere modificar la estructura
# del mensaje modificar el metodo de generar_reporte de GenerarInforme
generar_informes = GenerarInforme(archivo, hoja, guardar_txt=True)

# Diccionario que contiene el correo y el reporte a ser enviado
reportes = generar_informes.preparar_reportes_para_envio()

# Confirmación
print("Desea enviar los reportes mediante correo ? (S/n) ")
respuesta = input()
if respuesta == 'S':
    print("Enviando...")
    mail = Mail(asunto)
    mail.enviar_reportes(reportes)
else:
    print("NO se enviaron los reportes.")
    





