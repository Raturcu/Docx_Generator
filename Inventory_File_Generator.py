import datetime
from pathlib import Path
from tkinter import filedialog
import PySimpleGUI as sg
from docxtpl import DocxTemplate

#document_path = Path(__file__).parent / "your-document-template.docx"
document_path=Path("path-of-your-document")
doc = DocxTemplate(document_path)

today = datetime.datetime.today()
today_in_one_week = today + datetime.timedelta(days=7)
equip=["Desktop","Laptop","Workstation"]
model=["Optiplex 9020","Optiplex 7040","Optiplex 7050","Optiplex 7060","Optiplex 3060","Optiplex 5040","Optiplex 3080","Latitude 3520","Latitude 5420","Latitude 5520","Latitude 5591","Precision M3800","Elitebook 840","X1 Carbon"]
manuf=["Dell","HP","Lenovo"]
monitors=["0","1","2","3"]
layout = [
    [sg.Text("UTILIZATOR:"), sg.Input(key="UTILIZATOR", do_not_clear=False)],
    [sg.Text("PC NAME:"), sg.Input(key="PC_NAME", do_not_clear=False)],
    [sg.Text("Inventory Number:"), sg.Input(key="INVENTORY", do_not_clear=False)],
    [sg.Text("Serial Number:"), sg.Input(key="SERIAL", do_not_clear=False)],
    [sg.Text("TYPE:"), sg.DD(equip,key="TYPE",size=(43))],
    [sg.Text("MANUFACTURER:"), sg.DD(manuf,key="MANUF",size=(43))],
    [sg.Text("MODEL:"), sg.DD(model,key="MODEL",size=(43))],
    [sg.Text("SCREENS:"), sg.DD(monitors,key="NR",size=(43))],
    [sg.Button("Create Inventory File"), sg.Exit()],
    
]

window = sg.Window("Inventory File", layout, element_justification="right")

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "Create Inventory File":
        values["TODAY"] = today.strftime("%d-%m-%Y")
        # Render the template, save new word document & inform user
        doc.render(values)
        output_path = Path("output-path") / f"{values['UTILIZATOR']}.docx"
        doc.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")
        

window.close()
