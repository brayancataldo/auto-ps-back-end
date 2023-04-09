from photoshopy import Photoshopy

art = Photoshopy()

ARQUIVO_PHOTOSHOP = "./src/template.psd"
name = input("Nome: ")
year_quantity = input("Anos: ")
art.get_individual_artwork(psd_path=ARQUIVO_PHOTOSHOP, filename=name, name=name, year_quantity=year_quantity)