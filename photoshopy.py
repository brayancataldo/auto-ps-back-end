import win32com.client
import os
from pathvalidate import sanitize_filename
import openpyxl
from datetime import date, timedelta, datetime
import math
import logging

DO_NOT_SAVE_CHANGES = 2
DEFAULT_JPEG_QUALITY = 10
SAVE_FOR_WEB = 2
DOCUMENT_TYPE_JPEG = 6

class Photoshopy:
    app = None
    psd_file = None

    def __init__(self ):
        self.app = win32com.client.Dispatch("Photoshop.Application")
        self.app.Visible = True

        self.log = logging.getLogger(__name__)
        self.log.debug("Photoshop started")

    def closePhotoshop(self):
        self.app.Quit()
        self.log.debug("Photoshop closed")

    def openPSD(self, filename):
        if os.path.isfile(filename) == False:
            print("arquivo psd não existe")
            self.log.error("File not found: {0}".format(filename))
            self.closePhotoshop()
            return False

        self.app.Open(filename)
        self.psd_file = self.app.Application.ActiveDocument

        self.log.debug("PSD opened: {0}".format(filename))
        return True

    def closePSD(self):
        if self.psd_file is None:
            self.log.error("PSD file is not opened")
            raise Exception(FileNotFoundError)

        self.app.Application.ActiveDocument.Close(DO_NOT_SAVE_CHANGES)
        self.log.debug("PSD file closed")

    def updateLayerText(self, layer_name, text):
        if self.psd_file is None:
            self.log.error("PSD file is not opened")
            raise Exception(FileNotFoundError)
        
        layer = self.psd_file.ArtLayers[layer_name]
        layer_text = layer.TextItem
        layer_text.contents = text
        return True

    def exportJPEG(self, filename, folder='', quality=DEFAULT_JPEG_QUALITY):
        if self.psd_file is None:
            self.log.error("PSD file is not opened")
            raise Exception(FileNotFoundError)

        filename = sanitize_filename(filename)
        full_path = os.path.join(folder, filename)

        options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")
        options.Format = DOCUMENT_TYPE_JPEG
        options.Quality = quality

        self.psd_file.Export(ExportIn=full_path, ExportAs=SAVE_FOR_WEB, Options=options)
        self.log.info("JPEG Export: {0}".format(full_path))
        print("Imagem exportada com sucesso!")
        return os.path.isfile(full_path)


    def readTable(self, path):
        wb = openpyxl.load_workbook(filename=path)
        print('Lendo tabela.')
        array = []
        for d in wb['Página1'].iter_rows(values_only = True):
            if isinstance(d[1], datetime):
                date1 = min(datetime(2023, 1, 1, 0, 0), d[1])
                date2 = max(datetime(2023, 1, 1, 0, 0), d[1])
                delta = date2 - date1
                time_in_years = delta.days / 365.25
                rounded_years = math.ceil(time_in_years)
                array.append({"name": d[0], "year": rounded_years})
        return array
    
    def get_individual_artwork(self, psd_path, filename, name, year_quantity):
            print("Arte " + name + " em andamento.")
            self.jpeg_path = os.path.abspath('./src/export')
            year_word = "ANOS" if int(year_quantity) > 1 else "ANO"
            # Verifica se o arquivo PSD pode ser aberto
            if not self.openPSD(os.path.abspath(psd_path)):
                return

            # Atualiza o texto em três camadas diferentes
            if self.updateLayerText("YEAR_QUANTITY", year_quantity) and \
            self.updateLayerText("YEAR_WORD", year_word) and \
            self.updateLayerText("NAME", name.upper()):
                
                # Exporta o arquivo JPEG
                self.exportJPEG(name+".jpg", self.jpeg_path)

                # Exibe uma mensagem indicando que o processo foi concluído com sucesso
                print("Arte criada com sucesso!")
            
            # Fecha o arquivo PSD
            self.closePSD()

    def get_artwork_by_layers(self, new_name, psd_filename, layers):
        name, _ = psd_filename.split('.')
        psd_file="./src/import/" + psd_filename
        self.psd_origin = os.path.abspath(psd_file)
        self.jpeg_path = os.path.abspath("./src/export")
        self.jpeg_name = name + ".jpg"
        print("Working on " + name + " artwork")
        self.jpeg_full_path = os.path.join(self.jpeg_path, self.jpeg_name)

        if not os.path.exists(self.jpeg_path):
                    os.mkdir(self.jpeg_path)

        self.openPSD(self.psd_origin)
      
        opened = self.openPSD(self.psd_origin)
        if opened:
            for each in layers:
                self.updateLayerText(each['layerName'], each['value'])

            self.exportJPEG(new_name+".jpg", self.jpeg_path)
            self.closePSD()
        print("Done")





    

