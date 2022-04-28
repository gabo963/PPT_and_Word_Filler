from tkinter import *
import tkinter.filedialog

#La Ruta Inicial

cargados = 0

rutaExcel = ""
rutaWord = ""
rutaPPT = "" 

def bc1( main, me ):
    
    global rutaWord
    global rutaExcel
    global rutaPPT
    
    tempdir = tkinter.filedialog
    tempdir = tempdir.askopenfilenames(parent=main, filetypes=[("Excel files", "*.xlsx"),("Excel files", "*.xlsm") ], title='Seleccione su archivo')
    
    
    if tempdir != "":
        cadena = tempdir[0]
        
        if len(cadena) > 0:
            rutaExcel = cadena
            me['text'] = cadena[-(len(cadena)-cadena.rindex('/')-1):]
    else:
        me['text'] = 'Seleccionar Un Archivo'
        rutaExcel = ''
        
def bc2( main, me ):
    
    global rutaWord
    global rutaExcel
    global rutaPPT
    
    tempdir = tkinter.filedialog
    tempdir = tempdir.askopenfilenames(parent=main, filetypes=[("Word files", "*.docx")], title='Seleccione su archivo')
    
    
    if tempdir != "":
        cadena = tempdir[0]
        
        if len(cadena) > 0:
            rutaWord = cadena
            me['text'] = cadena[-(len(cadena)-cadena.rindex('/')-1):]
    else:
        me['text'] = 'Seleccionar Un Archivo'
        rutaWord = ''
    

def bc3( main, me ):
    
    global rutaWord
    global rutaExcel
    global rutaPPT
    
    tempdir = tkinter.filedialog
    tempdir = tempdir.askopenfilenames(parent=main, filetypes=[("PowerPoint files", "*.pptx") ], title='Seleccione su archivo')
    
    
    if tempdir != "":
        cadena = tempdir[0]
        
        if len(cadena) > 0:
            rutaPPT = cadena
            me['text'] = cadena[-(len(cadena)-cadena.rindex('/')-1):]
    else:
        me['text'] = 'Seleccionar Un Archivo'
        rutaPPT = ''
        