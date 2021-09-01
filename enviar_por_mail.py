import smtplib, ssl
import parametros as p
from time import sleep, time
from email.mime.text import MIMEText
from email.header import Header

class Mail:

    def __init__(self, asunto):
        self.port = 465
        self.smtp_server_domain_name = "smtp.gmail.com"  # Servidor de correos uc
        self.sender_mail = p.CORREO
        self.password = p.PASSWORD
        self.asunto = asunto

    def send(self, emails, subject, content):
        ssl_context = ssl.create_default_context()
        service = smtplib.SMTP_SSL(self.smtp_server_domain_name, self.port, context=ssl_context)
        service.login(self.sender_mail, self.password)
        
 

        for email in emails:
            # Mensaje
            mensaje = MIMEText(content, 'mixed', 'utf-8')
            mensaje['From'] = self.sender_mail
            mensaje['To'] = email
            mensaje['Subject'] = Header(self.asunto, 'utf-8')    

            result = service.sendmail(self.sender_mail, email, mensaje.as_string())

        service.quit()

    def enviar_reportes(self, reportes):
        # enviar los reportes a cada uno de los correos
        total_a_enviar = len(reportes)
        t = time()
        for correo_index, (correo, reporte) in enumerate(reportes.items()):
            print(f"Progreso {100*correo_index/total_a_enviar:0.1f}%, enviando a {correo}")
            # enviar

            #reporte = reporte.format(table=tabulate(reporte))
            
            self.send([correo], self.asunto, reporte)
            sleep(15)  # Delay entre correos no bajar demasiado o puede que se sature
            

        print("Todos los reportes fueron enviados de manera correcta :)")
        print(f"Se demoro {time() - t} segundos.")
            


    
