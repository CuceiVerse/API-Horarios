from selenium import webdriver
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import pandas as pd
import time
import os
from datetime import datetime


def obtener_ciclo_actual() -> str:
    """Determina el ciclo escolar actual de la UdeG (ej. '202610')."""
    hoy = datetime.now()
    anio = hoy.year
    mes = hoy.month
    # Calendario A (10): Enero–Junio | Calendario B (20): Julio–Diciembre
    sufijo = "10" if mes <= 6 else "20"
    return f"{anio}{sufijo}"


def scrape_oferta(ciclo: str = None, centro: str = "D") -> list[dict]:
    """
    Realiza el scraping de la oferta académica de SIIAU para el ciclo y centro dados.
    Retorna una lista de diccionarios con las materias encontradas.
    Si `ciclo` es None, usa el ciclo actual calculado automáticamente.
    """
    if ciclo is None:
        ciclo = obtener_ciclo_actual()

    ORDEN = "0"
    MOSTRAR = "500"

    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1280,2000")

    driver = webdriver.Chrome(options=opts)
    driver.get("http://consulta.siiau.udg.mx/wco/sspseca.forma_consulta")
    WebDriverWait(driver, 20)

    print(f"[Horarios API] Buscando oferta para ciclo={ciclo} centro={centro}...")

    Select(driver.find_element(By.NAME, "ciclop")).select_by_value(ciclo)
    Select(driver.find_element(By.NAME, "cup")).select_by_value(centro)
    driver.find_element(By.XPATH, f"//input[@name='ordenp'][@value='{ORDEN}']").click()
    driver.find_element(By.XPATH, f"//input[@name='mostrarp'][@value='{MOSTRAR}']").click()

    boton = driver.find_element(By.ID, "idConsultar")
    driver.execute_script("arguments[0].scrollIntoView();", boton)
    boton.click()

    datos = []
    pagina = 1

    while True:
        print(f"[Horarios API] Procesando página {pagina}...")
        time.sleep(4)

        soup = BeautifulSoup(driver.page_source, "html.parser")
        rows = soup.find_all("tr", style=lambda x: x and "background-color" in x)

        for row in rows:
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

                    ses_h_list, hora_list, dias_list, edif_list, aula_list, per_list = [], [], [], [], [], []
                    for tr_h in celdas[7].find_all("tr"):
                        tds = tr_h.find_all(["td", "th"])
                        if len(tds) >= 6:
                            ses_h_list.append(tds[0].get_text(strip=True))
                            hora_list.append(tds[1].get_text(strip=True))
                            dias_list.append(tds[2].get_text(strip=True))
                            edif_list.append(tds[3].get_text(strip=True))
                            aula_list.append(tds[4].get_text(strip=True))
                            per_list.append(tds[5].get_text(strip=True))

                    ses_p_list, prof_list = [], []
                    for tr_p in celdas[8].find_all("tr"):
                        tds = tr_p.find_all(["td", "th"])
                        if len(tds) >= 2:
                            ses_p_list.append(tds[0].get_text(strip=True))
                            prof_list.append(tds[1].get_text(strip=True))

                    datos.append({
                        "NRC": nrc,
                        "Clave": clave,
                        "Materia": materia,
                        "Sec": sec,
                        "CR": int(cr) if cr.isdigit() else cr,
                        "CUP": cup,
                        "DIS": dis,
                        "Hora": "\n".join(hora_list),
                        "Dias": "\n".join(dias_list),
                        "Edificio": "\n".join(edif_list),
                        "Aula": "\n".join(aula_list),
                        "Periodo": "\n".join(per_list),
                        "Profesor": "\n".join(prof_list),
                        # Raw arrays for richer clients
                        "sesiones": [
                            {
                                "ses": ses_h_list[i] if i < len(ses_h_list) else "",
                                "hora": hora_list[i] if i < len(hora_list) else "",
                                "dias": dias_list[i] if i < len(dias_list) else "",
                                "edificio": edif_list[i] if i < len(edif_list) else "",
                                "aula": aula_list[i] if i < len(aula_list) else "",
                                "periodo": per_list[i] if i < len(per_list) else "",
                            }
                            for i in range(len(ses_h_list))
                        ],
                    })
                except Exception as e:
                    print(f"[Horarios API] Advertencia fila: {e}")
                    continue

        siguientes = driver.find_elements(By.XPATH, "//input[contains(@value, '500 Pr')]")
        if siguientes:
            try:
                driver.execute_script("arguments[0].scrollIntoView();", siguientes[0])
                siguientes[0].click()
                pagina += 1
            except Exception as e:
                print(f"[Horarios API] Error paginacion: {e}")
                break
        else:
            break

    driver.quit()
    print(f"[Horarios API] Scraping finalizado. {len(datos)} registros en {pagina} páginas.")
    return datos


def exportar_excel(datos: list[dict], ciclo: str, centro: str) -> str:
    """Exporta los datos a un archivo Excel y retorna la ruta del archivo generado."""
    columnas = [
        "NRC", "Clave", "Materia", "Sec", "CR", "CUP", "DIS",
        "Hora", "Dias", "Edificio", "Aula", "Periodo", "Profesor",
    ]
    df = pd.DataFrame(datos, columns=columnas)

    directorio_actual = os.path.dirname(os.path.abspath(__file__))
    nombre_archivo = f"oferta_siiau_{ciclo}_{centro}.xlsx"
    archivo = os.path.join(directorio_actual, nombre_archivo)

    with pd.ExcelWriter(archivo, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Oferta")
        worksheet = writer.sheets["Oferta"]
        for column_cells in worksheet.columns:
            length = max(len(str(cell.value) if cell.value is not None else "") for cell in column_cells)
            adjusted_width = min(length + 2, 50)
            worksheet.column_dimensions[column_cells[0].column_letter].width = adjusted_width

    print(f"[Horarios API] Excel exportado: {archivo} ({len(df)} registros)")
    return archivo


# Si se ejecuta directamente (modo script legacy), hace el scraping y exporta Excel
if __name__ == "__main__":
    ciclo = obtener_ciclo_actual()
    centro = "D"
    datos = scrape_oferta(ciclo, centro)
    exportar_excel(datos, ciclo, centro)
