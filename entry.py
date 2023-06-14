import PySimpleGUI as sg
import pandas as pd

sg.theme('LightBlue')

EXCEL_FILE = 'Pendaftaran (1).xlsx'

df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Masukan Data Kamu: ')],
    [sg.Text('Nama', size=(15, 1)), sg.InputText(key='Nama')],
    [sg.Text('jabatan', size=(15, 1)), sg.InputText(key='Jabatan')],
    [sg.Text('divisi', size=(15, 1)), sg.InputText(key='Divisi')],
    [sg.Text('Tanggal mulai', size=(15, 1)), sg.InputText(key='TGL Mulai')],
    [sg.Text('Jenis surat', size=(15, 1)), sg.Combo(['izin', 'sakit', 'keperluan keluarga'], key='Jenis Surat')],
    [sg.Submit(), sg.Button('clear'), sg.Exit()]
]

window = sg.Window('Form Surat Staff AKPAR NHI Bandung', layout)

def clear_input():
    for key in values:
        window[key].update('')
    return None

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == 'Exit':
        break
    if event == 'clear':
        clear_input()
    if event == 'Submit':
        df = pd.concat([df, pd.DataFrame([values], columns=df.columns)], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data Berhasil Di Simpan')
        clear_input()

window.close()
