import PySimpleGUI as sg
from CsvExporter import CsvExporter


class CsvExporterGui:
    def __init__(self):
        sg.theme("SystemDefault")
        font = ("verdana, 12")
        layout = [
            [sg.T("")], [sg.Text("Escolha o Arquivo:", font=font, size=(15, 1)), sg.Input(
                key="caminho", enable_events=True), sg.FileBrowse()],
            [sg.Text("Escolha a Aba:", font=font, size=(15, 1)), sg.Input(
                key="aba", size=(30, 1))],
            [sg.Text("Nome do Diretório:", font=font, size=(15, 1)), sg.Input(
                key="diretorio", size=(30, 1))],
            [sg.Text("Colunas da Planilha:", font=font,
                     size=(15, 1)), sg.Output(key='_output_')],
            [sg.Text("Semana:", font=font), sg.Spin(values=["atual", "W01", "W02", "W03", "W04", "W05", "W06", "W07", "W08", "W09",
                                                            "W10", "W11", "W12", "W13", "W14", "W15", "W16", "W17", "W18",
                                                            "W19", "W20", "W21", "W22", "W23", "W24", "W25", "W26", "W27",
                                                            "W28", "W29", "W30", "W31", "W32", "W33", "W34", "W35", "W36",
                                                            "W37", "W38", "W39", "W40", "W41", "W42", "W43", "W44", "W45",
                                                            "W46", "W47", "W48", "W49", "W50", "W51", "W52", "W53"], key='semana')],
            [sg.Button("Enviar dados", key='btnSubmit', font=font),
             sg.Text("", font=font, key="resultado")]
        ]

        self.janela = sg.Window("CSV EXPORTER", layout,
                                size=(600, 260), auto_size_buttons=True, resizable=False,
                                auto_size_text=True)

    def iniciar(self):

        while True:
            self.event, self.values = self.janela.Read()
            if self.event == "caminho":
                dados = CsvExporter.carregarAbas(self.values['caminho'])
                for x in dados:
                    print(x)
            elif self.event == sg.WIN_CLOSED or self.event == 'Quit':
                break
            else:
                ce = CsvExporter(
                    self.values['caminho'], self.values['aba'], self.values['semana'], self.values['diretorio'])

                resultado = ce.exportarCSV()

                if(resultado):
                    self.janela['caminho'].update('')
                    self.janela['aba'].update('')
                    self.janela['diretorio'].update('')
                    self.janela['_output_'].update('')
                    self.janela['resultado'].update(
                        "Arquivo Criado com Sucesso!")
                else:
                    self.janela['caminho'].update('')
                    self.janela['aba'].update('')
                    self.janela['diretorio'].update('')
                    self.janela['_output_'].update('')
                    self.janela['resultado'].update("Erro ao criar arquivo!")


tela = CsvExporterGui()
tela.iniciar()
