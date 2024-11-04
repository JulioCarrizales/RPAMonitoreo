from selenium import webdriver
from selenium.webdriver.edge.options import Options

def abrir_pagina_en_edge(url):
    edge_options = Options()
    edge_options.add_argument("--ignore-certificate-errors")  # Ignora los errores de certificado
    edge_options.add_argument("--ignore-ssl-errors")  # Ignora errores SSL
    edge_options.add_argument("--start-maximized")  # Inicia en pantalla completa
    edge_options.add_argument("--inprivate")  # Modo InPrivate para evitar problemas de autenticación
    edge_options.add_argument("--no-sandbox")  # Evita problemas de sandboxing
    edge_options.add_argument("--disable-features=IsolateOrigins,site-per-process")  # Desactiva características que pueden causar errores

    # Crea el controlador de Edge con las opciones configuradas
    driver = webdriver.Edge(options=edge_options)
    driver.get(url)

# Llamada a la función con tu URL
abrir_pagina_en_edge("https://tsps-vip.bn.com.pe:8043")
