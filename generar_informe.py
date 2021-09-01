
from openpyxl import load_workbook
import os


class GenerarInforme:

    def __init__(self, archivo, hoja, guardar_txt=False):
        self.workbook = load_workbook(archivo, data_only=True)
        self.worksheet = self.workbook[hoja]
        self.guardar_txt = guardar_txt



    def generar_reporte(self):
        # Letra ~ indica la columna del excel en la cual extraer la información
        # Entrega el valor y comentario en formato de tupla de una celda
        
        reporte = f"""
        Hola {self.valor('D')} {self.valor('C')}!

        Items                                               nota      Comentario:
        Informe
            Calidad                                         {self.valor('E'):0.2f}      {self.comentario('E')}
            Explicaciones                               {self.valor('F'):0.2f}      {self.comentario('F')}   
            Redacción                                    {self.valor('G'):0.2f}      {self.comentario('G')}
            Ortografía                                     {self.valor('H'):0.2f}      {self.comentario('H')}
            Final                                              {self.valor('I'):0.2f}      {self.comentario('I')}



        """

        if self.guardar_txt:
            if not os.path.isdir('resumen'):
                os.mkdir('resumen')
            with open(f"resumen/{self.obtener_correo()}.txt", "w") as file:
                file.write(reporte)


        return reporte

    def valor(self, letra):
        # Entrega el valor de la celda
        fila = self.fila_actual  # self.fila_actual indica el alumno en el que se esta trabajando
        letra = letra.upper()
        return self.worksheet[f"{letra}{fila}"].value

    def comentario(self, letra):
        # Entrega el comementario de la celda
        fila = self.fila_actual # self.fila_actual indica el alumno en el que se esta trabajando
        letra = letra.upper()

        comentario = self.worksheet[f"{letra}{fila}"].comment
        if comentario is None:
            return '-'
        else:
            return str(comentario)

    def obtener_correo(self):
        # Se asume que el correo siempre estara en la primera columna
        columna_correo = 'A'

        return self.valor(columna_correo)
        


    def preparar_reportes_para_envio(self):
        # Se asume que la primera fila es el header y por lo tanto se ignora siempre
        # Por lo que se empezara a buscar alumnos desde la fila 2

        # Se itera sobre un numero grande para abarcar todos los alumnos
        # pero solo se consideraran los que tengan correo
        reportes = {}  # Reportes es un diccionario que contendra la informacion para enviar
        for index_alumno in range(2,2000): 
            # Obtener correo
            self.fila_actual = index_alumno
            correo = self.obtener_correo()

            if correo is None:
                continue
            
            # generar_reporte
            reporte = self.generar_reporte()



            # Guarda info para ser enviada
            reportes[correo] = reporte
            

            # siguiente alumno
            self.fila_actual +=1


        return reportes


            
            





        

        
        




if __name__ == '__main__':


    archivo = "notas.xlsx"
    hoja = "T1_mail"

    generar_informes = GenerarInforme(archivo, hoja, guardar_txt=True)
    generar_informes.preparar_reportes_para_envio()
    

