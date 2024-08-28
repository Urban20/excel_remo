from selenium import webdriver 
from bs4 import BeautifulSoup
import pandas,os,time,datetime,json

def personalizar_excel(c_encabezado,c_letras,filtro,data_frame,editor_,sheet_name):
    workbook  = editor_.book
    worksheet = editor_.sheets[sheet_name]
    header_format = workbook.add_format({
                    'bold': True,
                    'font_color': c_letras,
                    'bg_color': c_encabezado,
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter',
                    'font_size': 14 
                })
    
    cell_format = workbook.add_format({
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter',
                    
                })
    
    for col_num, value in enumerate(data_frame.columns.values):
        worksheet.write(0, col_num, value, header_format)
        
    for row_num in range(1, len(data_frame) + 1):
        for col_num in range(len(data_frame.columns)):
                        worksheet.write(row_num, col_num, data_frame.iloc[row_num - 1, col_num], cell_format)

        for col_num, value in enumerate(data_frame.columns):
                
            column_width = max(data_frame[value].astype(str).map(len).max(), len(value))
                
            worksheet.set_column(col_num, col_num, column_width + 10)  

            if filtro:  
                worksheet.autofilter(0, 0, len(data_frame), len(data_frame.columns) - 1)   


if __name__ == '__main__':
    
    carpeta = 'registros'
    fecha = str(datetime.datetime.now())[:10]
    os.chdir(f'{os.environ.get('USERPROFILE')}/Desktop')
    arch_json= 'data_remeros.json'
    try:
        with open(arch_json,'r') as arch:
            data_json=json.load(arch)['remeros']

    except FileNotFoundError:
        with open(arch_json,'w') as arch:
            data = {
                 'remeros':[]
            }
            json.dump(data,arch)
        
    
    os.makedirs(carpeta,exist_ok=True)
    os.chdir(f'{os.environ.get('USERPROFILE')}/Desktop/{carpeta}')

    try:
        while True:
            
            usuarios=sorted(data_json)
            
            nombre_remeros = []
            tiempo_remeros = []
            distancia = []

            opciones= webdriver.EdgeOptions()
            opciones.add_argument("--headless")
            opciones.add_argument("--disable-gpu")
            opciones.add_argument("--blink-settings=imagesEnabled=false")
                
            driver = webdriver.Edge(options=opciones)

            
            for usuario in usuarios:
                try:  
                    driver.get(f'https://www.relive.cc/profile/{usuario}')
                    time.sleep(2.5)

                    html= BeautifulSoup(driver.page_source,'html.parser')
                    contenedor= html.find('section',class_='d-flex flex-column overview-activity-list-container')
                
                    stats= contenedor.find_all('p',class_='text-lead',limit=2)
                            
                    remero= html.find('h3',class_='profile-name text-big').get_text()
                    nombre_remeros.append(remero)
                    #tiempo
                    tiempo_remeros.append(stats[0].get_text())
                    #distancia
                    distancia.append(stats[1].get_text())


                except AttributeError:
                                
                    nombre_remeros.append(html.find('h3',class_='profile-name text-big text-center').get_text())
                    tiempo_remeros.append('no encontrado')
                    distancia.append('no encontrado')
                except Exception as error:
                    print(f'error: {error}')
                    continue
                        
                

            try:
                excel = pandas.DataFrame(
                        {
                            'Remeros': nombre_remeros,
                            'Tiempos' : tiempo_remeros,
                            'Distancias': distancia

                        })
                with pandas.ExcelWriter(f'{fecha}.xlsx',engine='xlsxwriter') as editor:
                    excel.to_excel(editor,index=False,sheet_name='Datos')
                        
                    personalizar_excel(c_encabezado='712eb8',c_letras='white',sheet_name='Datos',filtro=False,data_frame=excel,editor_=editor)
                

            except PermissionError:
                    
                continue
            except Exception as error:
                print(f'error: {error}')
                continue

            finally:
                    
                driver.quit()
                print('finalizado')
                    
                time.sleep(900)
    except NameError:
        os.system('msg * falta configurar el archivo .json')