import pandas as pd
import os
from datetime import datetime


class CsvExporter:
    def __init__(self, caminho, aba, semana, diretorio="CSV") -> None:
        self.caminho = str(caminho)
        self.diretorio = diretorio
        self.aba = aba
        self.semana = semana

    def criarDiretorio(self, diretorio="CSV"):
        try:
            os.mkdir(diretorio)
            print("Diretório", diretorio, "Criado com sucesso")
            return os.path.abspath(diretorio)
        except FileExistsError:
            return os.path.abspath(diretorio)

    @staticmethod
    def carregarAbas(arquivo):
        try:
            path = os.path.abspath(arquivo)
            df = pd.ExcelFile(path)
            return df.sheet_names
        except:
            return {"mensagem": "Falha ao carregar o arquivo !"}

    def exportarCSV(self):
        try:
            diretorio = self.criarDiretorio(self.diretorio)
            if (self.semana == "atual"):
                semana = datetime.now().strftime("%W")
                arquivo = f'W{semana}.csv'
            else:
                arquivo = f'{self.semana}.csv'

            path = os.path.abspath(self.caminho)
            df = pd.read_excel(path, self.aba)
            df.to_csv(f'{diretorio}/{arquivo}', index=False, date_format="%A, %d,%B,%Y", columns=["AD site", "Computer name", "Last userID",
                                                                                                  "Model", "S/N", "Build release", "Computer type", "Computer role",
                                                                                                  "Workstation last active time", "Last hardware scan"])
            return True
        except:
            return False
