from selenium import webdriver
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import pandas as pd
import time

# Configuración
CICLO = "202520"
CENTRO = "D"
ORDEN = "0"
MOSTRAR = "500"

# Iniciar navegador
opts = Options()
opts.add_argument("--headless=new") # Lo hacemos headless para que no estorbe la ventana
opts.add_argument("--no-sandbox")
opts.add_argument("--disable-gpu")
opts.add_argument("--window-size=1280,2000")
driver = webdriver.Chrome(options=opts)
driver.get("http://consulta.siiau.udg.mx/wco/sspseca.forma_consulta")
wait = WebDriverWait(driver, 20)

print(f"Buscando ofertas para el ciclo {CICLO} y centro {CENTRO}...")

# Llenar el formulario
Select(driver.find_element(By.NAME, "ciclop")).select_by_value(CICLO)
Select(driver.find_element(By.NAME, "cup")).select_by_value(CENTRO)
driver.find_element(By.XPATH, f"//input[@name='ordenp'][@value='{ORDEN}']").click()
driver.find_element(By.XPATH, f"//input[@name='mostrarp'][@value='{MOSTRAR}']").click()

# Clic en Consultar
boton = driver.find_element(By.ID, "idConsultar")
driver.execute_script("arguments[0].scrollIntoView();", boton)
boton.click()

# Bucle para recorrer todas las páginas usando el botón "500 Próximos"
datos = []
pagina = 1

while True:
    print(f"Procesando página {pagina}...")
    time.sleep(4)  # Esperar que cargue la tabla HTML

    # Parsear HTML actual con BeautifulSoup
    soup = BeautifulSoup(driver.page_source, "html.parser")
    
    # Buscar todas las filas de materias
    rows = soup.find_all("tr", style=lambda x: x and "background-color" in x)
    
    for row in rows:
        # Tomar todas las celdas directas de la fila
        celdas = row.find_all("td", recursive=False)
        if len(celdas) >= 9:
            try:
                nrc = celdas[0].get_text(strip=True)
                clave = celdas[1].get_text(strip=True)
                materia = celdas[2].get_text(strip=True)
                sec = celdas[3].get_text(strip=True)
                cr = celdas[4].get_text(strip=True)
                cup = celdas[5].get_text(strip=True)
                dis = celdas[6].get_text(strip=True)
                
                # Extraer Horario desglosado (Ses, Hora, Dias, Edificio, Aula, Periodo)
                # Como puede haber múltiples horarios, los unimos por saltos de línea PERO dentro de su columna respectiva
                ses_h_list, hora_list, dias_list, edif_list, aula_list, per_list = [], [], [], [], [], []
                
                for tr_h in celdas[7].find_all('tr'):
                    tds = tr_h.find_all(['td', 'th'])
                    if len(tds) >= 6:
                        ses_h_list.append(tds[0].get_text(strip=True))
                        hora_list.append(tds[1].get_text(strip=True))
                        dias_list.append(tds[2].get_text(strip=True))
                        edif_list.append(tds[3].get_text(strip=True))
                        aula_list.append(tds[4].get_text(strip=True))
                        per_list.append(tds[5].get_text(strip=True))

                ses_h = "\n".join(ses_h_list)
                hora = "\n".join(hora_list)
                dias = "\n".join(dias_list)
                edif = "\n".join(edif_list)
                aula = "\n".join(aula_list)
                per = "\n".join(per_list)
                
                # Extraer Profesor desglosado (Ses, Profesor)
                ses_p_list, prof_list = [], []
                for tr_p in celdas[8].find_all('tr'):
                    tds = tr_p.find_all(['td', 'th'])
                    if len(tds) >= 2:
                        ses_p_list.append(tds[0].get_text(strip=True))
                        prof_list.append(tds[1].get_text(strip=True))
                
                ses_p = "\n".join(ses_p_list)
                prof = "\n".join(prof_list)
                
                datos.append([
                    nrc, clave, materia, sec, cr, cup, dis, 
                    ses_h, hora, dias, edif, aula, per, 
                    ses_p, prof
                ])
            except Exception as e:
                print(f"Advertencia: Error procesando una fila: {e}")
                continue

    # Buscar el botón de siguientes 500 resultados
    # El valor del botón suele ser "500 Próximos" (con posibles variaciones de encoding)
    siguientes = driver.find_elements(By.XPATH, "//input[contains(@value, '500 Pr')]")
    
    if siguientes:
        try:
            driver.execute_script("arguments[0].scrollIntoView();", siguientes[0])
            siguientes[0].click()
            pagina += 1
        except Exception as e:
            print(f"No se pudo hacer clic en el botón Siguiente: {e}")
            break
    else:
        # No hay más botón de "500 Próximos", terminamos
        break

driver.quit()

# Guardar Excel con los nombres de columnas exactos y separados solicitados
columnas = [
    "NRC", "Clave", "Materia", "Sec", "CR", "CUP", "DIS", 
    "Ses (Materia)", "Hora", "Dias", "Edificio", "Aula", "Periodo", 
    "Ses (Profesor)", "Profesor"
]
df = pd.DataFrame(datos, columns=columnas)
archivo = f"oferta_siiau_{CICLO}_{CENTRO}.xlsx"

# Guardar a un archivo excel y ajustar ancho de columnas
with pd.ExcelWriter(archivo, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Oferta')
    worksheet = writer.sheets['Oferta']
    for column_cells in worksheet.columns:
        length = max(len(str(cell.value) if cell.value is not None else "") for cell in column_cells)
        # Dar un poco más de margen
        adjusted_width = (length + 2)
        # Limitar ancho máximo para no hacer columnas exageradamente grandes
        if adjusted_width > 50:
            adjusted_width = 50
        worksheet.column_dimensions[column_cells[0].column_letter].width = adjusted_width

print(f"Excel generado con auto-ajuste y {len(df)} registros en total: {archivo}")
