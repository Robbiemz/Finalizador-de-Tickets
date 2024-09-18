# Import the library
import asana
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import time
import re
import json
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from colorama import init, Fore, Back, Style
init(autoreset=True)


chrome_options = Options()
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--headless')

# global:

#Asana
asana.Client.DEFAULT_OPTIONS['page_size'] = 100
asana.Client.LOG_ASANA_CHANGE_WARNINGS = False
personal_access_token = '1/1164896830095517:ac94baa7920396a4091f56519861aea1'
Tickets_Open='1203304323517458'
Tickets_Close='1203304323517459'
Seccion_Planificadas='1160149016649414'
Seccion_en_Proceso='1160149016649410'
Seccion_Revision='1165092959982793'
Seccion_Finalizadas='1160149016649411'

#OsTicket
Pagina_Soporte = 'https://soporte.grupoleiros.com/scp/login.php'
urlredirect = "https://soporte.grupoleiros.com/scp/tickets.php?id="
User = 'soporte'
password = '16072018A'


driver = webdriver.Chrome(options=chrome_options)

while True:
    print(Fore.LIGHTMAGENTA_EX + "                    .::Nueva Consulta:" + "  " + datetime.today().strftime('%d-%m-%y %H:%M')+ "::.")    
    
    contador_pag_blucle = 1

    # Asana client
    client = asana.Client.access_token(personal_access_token)

    contador_true=0

    #Consulta de tareas planificadas y en proceso
    #for result2 in client.tasks.get_tasks_for_tag(Tickets_Open, {'param': 'value', 'param': 'value'}, opt_pretty=True):
    #tck_Planificadas = client.tasks.get_tasks_for_section(Seccion_Planificadas, {'param': 'value', 'param': 'value'}, opt_pretty=True, opt_fields=['gid','name'])
    #tck_Proceso = client.tasks.get_tasks_for_section(Seccion_en_Proceso, {'param': 'value', 'param': 'value'}, opt_pretty=True, opt_fields=['gid','name'])
    tck_Revision = client.tasks.get_tasks_for_section(Seccion_Revision, {'param': 'value', 'section': 'value'}, opt_pretty=True, opt_fields=['gid','name'])


    #El formato dict se convierte en str
    dict1 = json.dumps(list(tck_Revision))
    str_tck = json.loads(dict1)
    str_gid = json.loads(dict1)
    ocurrencias = str(str_tck).count('gid')

    driver.get("https://soporte.grupoleiros.com/scp/login.php")
    driver.find_element('id','name').send_keys(User)
    driver.find_element('id','pass').send_keys(password)

    driver.find_element('name','submit').click()
    driver.get(urlredirect)
    html_contador = BeautifulSoup(driver.page_source,'html.parser')   

    #Se le da la informacion correcta a los str_tck y str_gid
    for i in range(0, int(ocurrencias)):
        #Se formatea el string en formato de json para mostrar solo TCK
        str_tck[i] = re.sub(".*?#", "", str(str_tck[i]), count = 1)
        str_tck[i] = re.sub("n.*", "", str(str_tck[i]))
        str_tck[i] = str(str_tck[i])[:-1]

        #Se formatea el string en formato de json para mostrar solo GID
        str_gid[i] = re.sub(",.*", "", str(str_gid[i]))
        str_gid[i] = str(str_gid[i])[:-1]
        str_gid[i] = re.sub(".*'", "", str(str_gid[i]))
                
        #print('Este es el GID: ', str_gid[i])
        #print('---------------------------------------')
        #print('Este es el TCK: ', str_tck[i])
        #print('---------------------------------------')


    #Busqueda de tickets recorriendo todas las paginas para cerrar
    for t in range(0, int(ocurrencias)):
        Numero_ticket = re.sub(".*?-", "", str(str_tck[t]), count = 1)
        driver.get(urlredirect + Numero_ticket)
        html_validador = BeautifulSoup(driver.page_source,'html.parser')
        formatted_validador = str(html_validador)

        #time.sleep(3)
        try:
            try:
              driver.find_element(By.XPATH,'/html/body/div[1]/div[2]/div/div[4]/div/div[4]/form[1]/table/tbody[3]/tr[1]/td[2]/div[1]/ul/li[1]').click()
            except:
              driver.find_element(By.XPATH,'/html/body/div[1]/div[2]/div/div[4]/div/div[3]/form[1]/table/tbody[3]/tr[1]/td[2]/div[1]/ul/li[1]/a').click()   
            contador_comentario = 1
            none_body = client.stories.get_stories_for_task(str_gid[t], {'param': 'value', 'param': 'value'}, opt_pretty=True, opt_fields=['html_text'])
            none_body_str = str(list(none_body))
                        
                        
            if none_body_str.count('body') >= 1:
                for content_stories in client.stories.get_stories_for_task(str_gid[t], {'param': 'value', 'param': 'value'}, opt_pretty=True, opt_fields=['html_text']):
                    content_stories_html = content_stories['html_text']
                    #content_stories_formatter = content_stories_html.split('\n')
                    content_stories_body = BeautifulSoup(content_stories_html,"html.parser")
                    content_stories_find_body= content_stories_body.find('body')
                    if content_stories_find_body is not None:
                        content_stories_formatter_body = content_stories_find_body.get_text()
                            #Si no hay ningun comentario con etiqueta de @ para nombrar a una persona entonces procede a comentar normal
                        if str(content_stories_formatter_body).count('@') == 0:
                            driver.find_element(By.XPATH,'//*[@id="response"]').send_keys('comentario N°',contador_comentario,': ',content_stories_formatter_body,'<br>')
                            print(Fore.CYAN + 'COMENTARIO:', 'El', str_tck[t],'ha agregado el Comentario N°', contador_comentario,':',content_stories_formatter_body)
                            contador_comentario = contador_comentario + 1
                          
                            #Si hay comentario con etiqueta de @ para nombrar a una persona entonces se formatea el comentario para evitar eso
                        else:
                            content_stories_formatter_body_formatter = content_stories_formatter_body
                            content_stories_formatter_body_formatter = re.sub(".*? ", "", content_stories_formatter_body_formatter, count=2)
                            driver.find_element(By.XPATH,'//*[@id="response"]').send_keys('Comentario N°',contador_comentario,': ',content_stories_formatter_body_formatter,'<br>')   
                            print(Fore.CYAN + 'COMENTARIO:', 'El', str_tck[t],'ha agregado el Comentario N°', contador_comentario,':',content_stories_formatter_body_formatter)
                            contador_comentario = contador_comentario + 1
                          
                            #Si hay una imagen en el comentario se adjunta 
                        if none_body_str.count('img') >= 1:
                            content_stories_find_img= content_stories_body.find('img')
                            driver.find_element(By.XPATH,'//*[@id="response"]').send_keys('Imagen adjunta:',str(content_stories_find_img),'<br>')
                            print(Fore.CYAN + 'COMENTARIO:', ' Se ha adjuntado una imagen')
            else:
                driver.find_element(By.XPATH,'//*[@id="response"]').send_keys('comentario N° ',contador_comentario,': .::Listo, Ticket Resuelto::. <br>')
                print(Fore.CYAN + 'COMENTARIO:', 'El', str_tck[t],'ha agregado el Comentario N°', contador_comentario,':','Listo, no tenia comentario en asana')
                #time.sleep(5)        

            driver.find_element(By.XPATH,'//*[@id="resp_sec"]/tr[3]/td[2]/select').send_keys('Resuelto')
            print(Fore.GREEN + '   SUCCESS:','El', str_tck[t],'Se ha marcado como resuelto en OSticket')
            driver.find_element(By.XPATH, '//*[@id="reply"]/p/input[1]').click()
            print(Fore.GREEN + '   SUCCESS:','El', str_tck[t],'Se ha finalizado el ticket en Osticket')
            result = client.tasks.update_task(str_gid[t], {'completed': True,'assignee_section':Seccion_Finalizadas}, opt_pretty=True)                     
            print(Fore.GREEN + '   SUCCESS:','El', str_tck[t], 'esta asociado al GID:', str_gid[t], 'se ha finalizado satisfactoriamente')
            contador_true=1                                              
            #time.sleep(2)
            
        except:
                
            print(Fore.MAGENTA + '     ERROR:', 'El ' + str_tck[t] + ' tiene problemas para ser modificado')
            
    if contador_true==0:
        print(Fore.YELLOW + "                    .::No hay tickets por cerrar::.")   
                
    time.sleep(60)        
    