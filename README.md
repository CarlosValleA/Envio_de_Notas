# Envio de Notas
Script para enviar reportes de notas por correo de manera automatica.

El archivo principal a ejecutar es `main.py` y requiere de la libreria `openpyxl`, que se puede instalar con `pip install openpyxl`. Dada la configuración solo esta pensado para ser usado con cuentas google y requiere modificaciones para ser usados con servicios de mail distintos. 

## Pasos para utilizarlo

### 1. Agregar credenciales de correo en el script y autorizarlo
Es necesario completar las variables de `CORREO` y `PASSWORD` en el archivo `parametros.py`, con la información del mail con el que planea enviar los reportes de notas. Ademas, debe activar en el siguiente [link](https://myaccount.google.com/lesssecureapps) la utilización de aplicaciones menos seguras.


### 2. Preparar archivo excel
Para hacer uso del script, es necesario un archivo excel en donde las filas correspondan a los alumnos y las columnas los parametros a evaluar. Se puede utilizar la plantilla `notas.xlsx` como base. El script siempre ignora el header (la primera fila o fila 1), por lo ahi nunca debe haber correos y/o notas a enviar.

![Ejemplo notas.xlsx](https://i.imgur.com/c6jBOve.png)

**Es importante mencionar que el script asume que el correo se encuentra en la columna `A`**, mientras que todos los demas campos son opcionales. 


### 3. Preparar elementos en el correo
Cuando el archivo excel este preparado, es necesario anotar el nombre del archivo y la hoja en el que se encuentran las notas a enviar, para luego ser asignadas a las variables de `archivo` y `hoja` (strings) en `main.py`. Ademas, se debe indicar el `asunto` en formato de string.

### 4. Estructura del mensaje
La estructura del mesaje puede (y debe) ser modificada de acuerdo a cada necesidad. En el archivo `generar_informe.py` en el metodo de `generar_reporte` se puede crear la estructura del correo a enviar. Un ejemplo de mensaje es:


```python
mensaje = f"""

Hola {self.valor('D')}!,

En este correo se encuentra la evaluación de tu tarea.

Items                                       nota      Comentario:
Informe
    Calidad                                 {self.valor('E'):0.2f}      {self.comentario('E')}
    Explicaciones                           {self.valor('F'):0.2f}      {self.comentario('F')}   
    Redacción                               {self.valor('G'):0.2f}      {self.comentario('G')}
    Ortografía                              {self.valor('H'):0.2f}      {self.comentario('H')}
    Final                                   {self.valor('I'):0.2f}      {self.comentario('I')}

"""

```

El methodo `self.valor` obtiene la información que contiene la celda y `self.comentario` el comentario que contiene la celda. Solo es necesario indicar la letra de la columna, ya que se asume que en las filas se encuentran los alumnos y por lo tanto iterara sobre las filas para generar los reportes. Volver a mencionar que el correo del alumno siempre debe estar en la columna `A` y los demas campos son opcionales. Si el correo del alumno se encuentra vacio, dicho reporte sera ignorado y no se hara un intento de enviarlo, por lo cual sirve para ignorar enviar reportes a alumnos que ya no se encuetren en el ramo, sin tener que borrar o modificar filas.

###  5. Ejecutar main
Para comenzar la generación de los reportes y el envio de estos basta con ejectuar `main.py`. Para evitar envios de reportes por error, es necesario confirmar al mensaje `"Desea enviar los reportes mediante correo ? (S/n) "` con un `S` .


## Recomendaciones
En `main.py` en la clase de `GenerarInforme` es posible indicar la opción de `guardar_txt` como `True` o `False`. Cuando `guardar_txt=True` se guardaran archivos .txt con el mensaje a enviar a cada correo. Se recomienda hacer esto antes de enviar los correos, para comprar que el mensaje se encuentra correcto en su contenido. 



## Errores conocidos
Por temas de encoding, es necesario utilizar `ascii` en lugar de `utf-8`, lo que genera que los tabs no se representen de manera correcta (o se ignoren (?)), por lo que el mensaje enviado puede que no sea igual en cuanto a tabs que el .txt. Por lo que se necesita ajustar los tabs con espacios en `generar_reporte` mencionado en el paso 4.

Ejemplo

![](https://i.imgur.com/73R0Gls.png)
**A. Vista en .txt , B. Vista en el correo**

