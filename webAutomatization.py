from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.chrome.options import Options as ChromeOptions
import time
import pandas as pd
import os
from pathlib import Path


def download_PDF(excel_file, obj, obj_percent):
    df = pd.read_excel(
        excel_file,
        engine="openpyxl",
        header=None,
        names=["cliente", "conta", "empresa", "regiao", "assessor"],
    )
    lista_contas_btg = df["conta"].unique().tolist()
    options = ChromeOptions()
    # Get the user's home directory
    home_dir = os.path.expanduser("~")

    # Construct the path to the Downloads folder
    download_folder = os.path.join(home_dir, "Downloads")
    btg_folder = os.path.join(download_folder, "planilhas_btg")

    # check if directory exists
    if not os.path.exists(btg_folder):
        # if not, create the directory
        os.makedirs(btg_folder)

    prefs = {
        "download.default_directory": btg_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
    }
    options.add_experimental_option("prefs", prefs)
    # Get url page
    driver = webdriver.Chrome("chromedriver.exe", options=options)
    driver.maximize_window()
    driver.get("https://access.btgpactualdigital.com/login/externo")

    # Find the login and password inputs and insert values into them

    driver.find_element(
        By.CSS_SELECTOR,
        "div.authenticate-panel__form--content:nth-child(1) > input:nth-child(1)",
    ).send_keys("01005818118")
    driver.find_element(
        By.CSS_SELECTOR,
        "div.authenticate-panel__form--content:nth-child(2) > input:nth-child(1)",
    ).send_keys("001122")

    # Find the operation button to lead to search for the clients reports
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located(
            (
                By.ID,
                "operation_button",
            )
        )
    ).click()

    # Find the "atendimento ao cliente" button to open the search bar
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located(
            (
                By.ID,
                "feature[object Object]1_button",
            )
        )
    ).click()

    # Loop for download both reports from each account
    # looks for the search bar, inserts the acount id and click enter
    bar = 0
    for cliente in lista_contas_btg:
        bar += 0.5
        obj.UpdateBar(bar, len(lista_contas_btg))
        obj_percent.update(str(int((bar / len(lista_contas_btg)) * 100)) + "%")

        # Clears the inputbox if exists something
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (
                    By.CSS_SELECTOR,
                    ".form-control",
                )
            )
        ).clear()
        # inserts the account id on the inputbox
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (
                    By.CSS_SELECTOR,
                    ".form-control",
                )
            )
        ).send_keys(cliente)
        # searches for the dropdown item of the account and clicks it
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    ".dropdown-item",
                )
            )
        ).click()

        # finds the element "histórico" to open the hovermenu
        historico = driver.find_element(
            By.CSS_SELECTOR,
            "#menu-list-history",
        )

        # Moves the mouse to the "histórico" and holds it there to open the hovermenu options
        hover = ActionChains(driver).move_to_element(historico)
        hover.perform()

        # wait to click the documents button inside the hover menu
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable(
                (
                    By.ID,
                    "submenu-documents",
                )
            )
        ).click()

        # wait to click the "Relatório de Perfomance" button, on the hovermenu
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    "#submenu-documents-report-performance > redirect-report-performance",
                )
            )
        ).click()

        # Saves the tab of the BTG url to further changing in tabs
        btg_tab = driver.current_window_handle

        # Waits 15 seconds to avoid BTG's autenthication
        time.sleep(15)

        # Switches to the download tab
        driver.switch_to.window(driver.window_handles[1])

        # clicks on the download button
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    "#builder > div > div.q-page-container.reset-padding > main > header > div > div > div.row.full-height > div.col-6.account-section__info.q-pl-xl > div.q-mt-xl > div > div:nth-child(2) > div > div:nth-child(2) > button",
                )
            )
        ).click()

        # Waitrs 10 seconds to get pass the authentication of the download
        time.sleep(15)

        # Close the second tab
        driver.close()

        # Switch back to the main tab
        driver.switch_to.window(btg_tab)

        # Moves again to the "histórico" button to get the other report file
        hover = ActionChains(driver).move_to_element(historico)
        hover.perform()

        # Clicks on the "Extrato da conta investimento" button
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable(
                (
                    By.ID,
                    "submenu-checking-account",
                )
            )
        ).click()

        # finds the element "histórico" to open the hovermenu
        filtrar = driver.find_element(
            By.ID,
            "product-exercicio-opcoes-exercido--button",
        )
        # Scroll to the button
        actions = ActionChains(driver)
        actions.move_to_element(filtrar).perform()

        # Clicks on the download button
        try:
            WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable(
                    (
                        By.ID,
                        "product-exercicio-opcoes-exercido--button",
                    )
                )
            ).click()
        except:
            continue
        time.sleep(5)

    driver.close()
    return btg_folder
