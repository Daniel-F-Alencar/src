import PySimpleGUI as sg
import pandas as pd
import threading
from webAutomatization import download_PDF
from upload_to_sharepoint import organize_reports


class MainApp:
    def __init__(self):
        # Theme
        sg.theme("darkblue17")

        # Layouts
        layout = [
            [
                sg.Text(
                    "Download de planilhas na plataforma BTG",
                    font="Arial 20",
                    key="1",
                    text_color="white",
                    size=(605, 2),
                    justification="center",
                )
            ],
            # Menu
            [sg.T(size=(20, 1), key="3")],
            [
                sg.Text(
                    "Selecione a planilha:",
                    justification="left",
                    size=(625),
                    font="Arial 10",
                    key="8",
                )
            ],
            [
                sg.Input(size=(70, 2), key="dir_files", background_color="white"),
                sg.FileBrowse(
                    "Pesquisar",
                    enable_events=True,
                    key="file",
                    file_types=(("xlsx", "*.xlsx"), ("csv", "*.csv")),
                    size=(50, 1),
                ),
            ],
            [sg.T(visible=True, key="10")],
            [sg.T(size=(20, 2), key="6")],
            [
                sg.Button("Processar", font="Arial 10 bold", size=(25, 2)),
                sg.T(size=(20, 2)),
                sg.Button("Resetar", font="Arial 10 bold", size=(25, 2), key="7"),
            ],
            [
                sg.Text(
                    "0%",
                    size=(625, 1),
                    font="Arial 10",
                    key="percent",
                    justification="center",
                )
            ],
            [
                sg.ProgressBar(
                    1,
                    orientation="h",
                    size=(605, 30),
                    key="progress",
                    pad=(5, 5),
                    bar_color=("green", "gray"),
                )
            ],
        ]

        # Create the Window
        self.window = sg.Window(
            "BTG webscraping", layout, size=(625, 290), icon=r"favicon.ico"
        )

    def download_and_organize_reports(self, values):
        folder = download_PDF(
            values["dir_files"],
            self.window.FindElement("progress"),
            self.window.FindElement("percent"),
        )
        df = pd.read_excel(
            values["dir_files"],
            engine="openpyxl",
            header=None,
            names=["cliente", "conta", "empresa", "regiao", "assessor"],
        )
        organize_reports(
            df,
            self.window.FindElement("progress"),
            self.window.FindElement("percent"),
        )

    def start_selenium_thread(self, values):
        thread = threading.Thread(
            target=self.download_and_organize_reports, args=(values,)
        )
        thread.daemon = True
        thread.start()

    def start(self):
        count = 0
        timeout1 = 10

        while True:
            # Update the PySimpleGUI window while the Selenium function is running
            event, values = self.window.read(timeout=timeout1)
            if count == 0:
                timeout1 = 50000000
                count = 1

            if event == sg.WINDOW_CLOSED:
                break

            # Reset
            if event == "7":
                self.window["dir_files"].update("")
                continue

            # Verification
            if values["dir_files"] == "" and event != "__TIMEOUT__":
                sg.popup("Arquivo não selecionado.", title="Atenção!")
                continue

            if event == "Processar":
                self.window.size = (625, 350)
                sg.popup(
                    """O download dos relatórios irá começar. 
Por favor, pegue o token em seu celular e aguarde!""",
                    title="Informativo",
                )
                self.window.refresh()
                self.start_selenium_thread(
                    values
                )  # Start Selenium in a separate thread

            # Restart
            if values["dir_files"] != "" and event != "Resetar" and event != "7":
                self.window.size = (625, 350)
                self.window["dir_files"].update("")
                self.window["progress"].update(0, 0)
                self.window["percent"].update("0%")


tela = MainApp()
tela.start()
