from photoshopy import Photoshopy

art = Photoshopy()

list = art.readTable("./src/import/table.xlsx")

for i, each in enumerate(list):
    names = each['name'].split()
    name = names[0] + " " + names[len(names) - 1] 
    art.get_individual_artwork(psd_path="./src/import/template.psd", filename=each['name'], name=name, year_quantity=str(each['year']))
