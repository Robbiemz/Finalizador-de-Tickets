from selenium.webdriver.chrome.options import Options
from keep_alive import keep_alive
from colorama import init, Fore, Back, Style
from selenium.webdriver.common.by import By
from selenium import webdriver
from datetime import datetime
from bs4 import BeautifulSoup
import asana
import json
import time
import re


init(autoreset=True)

chrome_options = Options()
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--headless')
driver = webdriver.Chrome(options=chrome_options)

# Asana
asana.Client.DEFAULT_OPTIONS['page_size'] = 100
asana.Client.LOG_ASANA_CHANGE_WARNINGS = False
personal_access_token = '1/1164896830095517:ac94baa7920396a4091f56519861aea1'
Tickets_Open = '1203304323517458'
Tickets_Close = '1203304323517459'
Seccion_Planificadas = '1160149016649414'
Seccion_en_Proceso = '1160149016649410'
Seccion_Revision = '1165092959982793'
Seccion_Finalizadas = '1160149016649411'

# OsTicket
Pagina_Soporte = 'https://soporte.grupoleiros.com/scp/login.php'
urlredirect = "https://soporte.grupoleiros.com/scp/tickets.php?id="
User = 'soporte'
password = '16072018A'
keep_alive()

# Inicializa Asana client
client = asana.Client.access_token(personal_access_token)

# Login OsTickets
driver.get("https://soporte.grupoleiros.com/scp/login.php")
driver.find_element('id', 'name').send_keys(User)
driver.find_element('id', 'pass').send_keys(password)
driver.find_element('name', 'submit').click()
driver.get(urlredirect)
html_contador = BeautifulSoup(driver.page_source, 'html.parser')



while True:
    print(Fore.LIGHTMAGENTA_EX + "MOMENTO DE EJECUCION", ":", datetime.today().strftime('%d-%m-%y | %H:%M'))

    # Resetear Contadores
    contador_pag_blucle = 1
    contador_true = 0
    
    # Consulta de tareas en Revision
    tck_Revision = client.tasks.get_tasks_for_section(Seccion_Revision, {'param': 'value', 'section': 'value'}, opt_pretty=True, opt_fields=['gid','name'])

    # Tareas en formato dict se convierte en str
    dict1 = json.dumps(list(tck_Revision))
    str_tck = json.loads(dict1)
    str_gid = json.loads(dict1)
    ocurrencias = len(str_tck)
    
    ## Formateo a los str_tck y str_gid
    #for i in range(0, int(ocurrencias)):
    #    # Se formatea el string en formato de json para mostrar solo TCK
    #    str_tck[i] = re.sub(".*?#", "", str(str_tck[i]), count=1)
    #    str_tck[i] = re.sub("\s+.*", "", str(str_tck[i]), count=1)   
#
    #    # Se formatea el string en formato de json para mostrar solo GID
    #    str_gid[i] = re.sub(",.*", "", str(str_gid[i]))
    #    str_gid[i] = re.sub(".*\s+", "", str(str_gid[i]))
    #    str_gid[i] = re.sub(".*:", "", str(str_gid[i]))
    #    str_gid[i] = re.sub(r'[^\w\s]', "", str(str_gid[i]))
        
    # Busqueda de tickets recorriendo todas las paginas para cerrar
    for t in range(0, int(ocurrencias)):
        
        str_tck[t] = re.sub(".*?#", "", str(str_tck[t]), count=1) # Se formatea el string en formato de json para mostrar solo TCK
        str_tck[t] = re.sub("\s+.*", "", str(str_tck[t]), count=1)  
        
        
        str_gid[t] = re.sub(",.*", "", str(str_gid[t])) # Se formatea el string en formato de json para mostrar solo GID
        str_gid[t] = re.sub(".*\s+", "", str(str_gid[t]))
        str_gid[t] = re.sub(".*:", "", str(str_gid[t]))
        str_gid[t] = re.sub(r'[^\w\s]', "", str(str_gid[t]))
        
        Numero_ticket = re.sub(".*?-", "", str(str_tck[t]), count=1) # Se formatea el sting de tickets para solo mostrar numeros
        
        driver.get(urlredirect + Numero_ticket) # Se abre la pagina del Ticket
        time.sleep(2)
        html_validador = BeautifulSoup(driver.page_source, 'html.parser') # Leer HTML
        formatted_validador = str(html_validador) # Formatear HTML a string

        try:
            try:
                driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[4]/div/div[4]/form[1]/table/tbody[3]/tr[1]/td[2]/div[1]/ul/li[1]').click()
            except:
                driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[4]/div/div[3]/form[1]/table/tbody[3]/tr[1]/td[2]/div[1]/ul/li[1]/a').click()
            
            contador_comentario = 1
            none_body = client.stories.get_stories_for_task(str_gid[t], {'param': 'value', 'param': 'value'}, opt_pretty=True, opt_fields=['html_text'])
            none_body_str = str(list(none_body))

            if none_body_str.count('body') >= 1:
                for content_stories in client.stories.get_stories_for_task(str_gid[t], {'param': 'value', 'param': 'value'}, opt_pretty=True, opt_fields=['html_text']):
                    content_stories_html = content_stories['html_text']
                    content_stories_body = BeautifulSoup(content_stories_html, "html.parser")
                    content_stories_find_body = content_stories_body.find('body')
                    
                    if content_stories_find_body is not None:
                        content_stories_formatter_body = content_stories_find_body.get_text()
                        
                        if str(content_stories_formatter_body).count('@') == 0:
                            driver.find_element(By.XPATH, '//*[@id="response"]').send_keys(
                                'comentario N°', contador_comentario, ': ', content_stories_formatter_body, '<br>')
                            print(Fore.CYAN + 'COMENTARIO:', 'El', str_tck[t], 'ha agregado el Comentario N°',
                                  contador_comentario, ':', content_stories_formatter_body)
                            contador_comentario += 1

                        else:
                            content_stories_formatter_body_formatter = content_stories_formatter_body
                            content_stories_formatter_body_formatter = re.sub(".*? ", "", content_stories_formatter_body_formatter, count=2)
                            driver.find_element(By.XPATH, '//*[@id="response"]').send_keys(
                                'Comentario N°', contador_comentario, ': ', content_stories_formatter_body_formatter, '<br>')
                            print(Fore.CYAN + 'COMENTARIO:', 'El', str_tck[t], 'ha agregado el Comentario N°',
                                  contador_comentario, ':', content_stories_formatter_body_formatter)
                            contador_comentario += 1

                        if none_body_str.count('img') >= 1:
                            content_stories_find_img = content_stories_body.find('img')
                            driver.find_element(By.XPATH, '//*[@id="response"]').send_keys(
                                'Imagen adjunta:', str(content_stories_find_img), '<br>')
                            print(Fore.CYAN + 'COMENTARIO:', ' Se ha adjuntado una imagen')
            else:
                driver.find_element(By.XPATH, '//*[@id="response"]').send_keys(
                    'comentario N° ', contador_comentario, ': .::Listo, Ticket Resuelto::. <br>')
                print(Fore.CYAN + 'COMENTARIO:', 'El', str_tck[t], 'ha agregado el Comentario N°', contador_comentario, ':',
                      'Listo, no tenía comentario en Asana')

            driver.find_element(By.XPATH, '//*[@id="resp_sec"]/tr[3]/td[2]/select').send_keys(
                'Resuelto')
            print(Fore.GREEN + 'SUCCESS:', 'El', str_tck[t], 'Se ha marcado como resuelto en OsTicket')
            driver.find_element(By.XPATH, '//*[@id="reply"]/p/input[1]').click()
            print(Fore.GREEN + 'SUCCESS:', 'El', str_tck[t], 'Se ha finalizado el ticket en OsTicket')
            result = client.tasks.update_task(str_gid[t], {'completed': True, 'assignee_section': Seccion_Finalizadas},
                                              opt_pretty=True)
            print(Fore.GREEN + 'SUCCESS:', 'El', str_tck[t], 'esta asociado al GID:', str_gid[t],
                  'se ha finalizado satisfactoriamente')
            contador_true = 1

        except Exception as e:
            #print(Fore.MAGENTA + 'ERROR:', 'El', str_tck[t], 'tiene problemas para ser modificado')
          print(Fore.MAGENTA + 'ERROR:', {e})

    if contador_true == 0:
        print(Fore.YELLOW + "                    .::No hay tickets por cerrar::.")

    time.sleep(60)
