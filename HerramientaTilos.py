from tkinter import *
import tkinter.filedialog
import os

import ComandosInterfaz

# Trae el mundo

import ExtraccionWORD
import ExtraccionPPT
import UbicarWORD
import UbicarPPT

if __name__ == "__main__":
    encendido = True

    tk = Tk()
    tk.title( 'Remplaza Texto' )
    tk.iconbitmap('logo_tilos.ico')
    tk.resizable(0, 0)

    # Titulo inicial

    def check( rutaWord, rutaExcel, rutaPPT ):

        global encendido
        global b4
        global b5
        
        b4['text'] = 'Extraer los parámetros de word.'
        b5['text'] = 'Extraer los parámetros de PowerPoint.'
        b6['text'] = 'Ubicar los parámetros en el word.'
        b7['text'] = 'Ubicar los parámetros en el PowerPoint.'

        if encendido:
            b4['state'] = 'disabled'
            b5['state'] = 'disabled'
            
            if rutaWord != "" and rutaExcel != "":
            
                b6['state'] = 'normal'
            else:
                b6['state'] = 'disabled'
            
            if rutaPPT != "" and rutaExcel != "":
            
                b7['state'] = 'normal'
            else:
                b7['state'] = 'disabled'
            
        else:
            b6['state'] = 'disabled'
            b7['state'] = 'disabled'
            
            if rutaWord != "" and rutaExcel != "":
            
                b4['state'] = 'normal'
            else:
                b4['state'] = 'disabled'
            
            if rutaPPT != "" and rutaExcel != "":
            
                b5['state'] = 'normal'
            else:
                b5['state'] = 'disabled'
            
        pass

    def switch( ):
        global encendido
        global b0
        
        encendido = not encendido
        
        if encendido:
            b0['text'] = 'Llenar Documentos.'
        else:
            b0['text'] = 'Extraer Parámetros.'
            
        check(ComandosInterfaz.rutaWord, ComandosInterfaz.rutaExcel, ComandosInterfaz.rutaPPT)
        
        pass

    def word(b2):
        
        ComandosInterfaz.bc2(tk,b2)
        check(ComandosInterfaz.rutaWord, ComandosInterfaz.rutaExcel, ComandosInterfaz.rutaPPT)
        pass

    def ppt(b3):
        
        ComandosInterfaz.bc3(tk,b3)
        
        check(ComandosInterfaz.rutaWord, ComandosInterfaz.rutaExcel, ComandosInterfaz.rutaPPT)
        pass

    def excel(b1):
        
        ComandosInterfaz.bc1(tk,b1)
        
        check(ComandosInterfaz.rutaWord, ComandosInterfaz.rutaExcel, ComandosInterfaz.rutaPPT)
        pass

    def ubicarWord(b6):
        try:
            UbicarWORD.ubicarPalabraW(ComandosInterfaz.rutaWord, ComandosInterfaz.rutaExcel, b6)
            
            mensaje = 'Word creado en la misma carpeta del archivo seleccionado.'
            tkinter.messagebox.showinfo( "Completado" , mensaje )
        except Exception as e:
            mensaje = 'No fue posible completar la operación.'
            tkinter.messagebox.showwarning( "Error" , mensaje )
            
            tkinter.messagebox.showwarning( "Error" , e.__class__ )

    def ubicarPpt(b7):
        try:
            UbicarPPT.ubicarPalabraP(ComandosInterfaz.rutaPPT, ComandosInterfaz.rutaExcel, b7)

            mensaje = 'PowerPoint creado en la misma carpeta del archivo seleccionado.'
            tkinter.messagebox.showinfo( "Completado" , mensaje )
        except Exception as e:
            mensaje = 'No fue posible completar la operación.'
            tkinter.messagebox.showwarning( "Error" , mensaje )
            
            tkinter.messagebox.showwarning( "Error" , e.__class__ )

    def extraerWord(b4):
        try:
            ExtraccionWORD.agregar_a_excelW(ComandosInterfaz.rutaWord, ComandosInterfaz.rutaExcel, b4)

            mensaje = 'Datos del Word exportados a Excel.'
            tkinter.messagebox.showinfo( "Completado" , mensaje )
        except Exception as e:
            mensaje = 'No fue posible completar la operación.'
            tkinter.messagebox.showwarning( "Error" , mensaje )
            
            tkinter.messagebox.showwarning( "Error" , e.__class__ )

            
    def extraerPpt(b5):
        try:
            ExtraccionPPT.agregar_a_excelP(ComandosInterfaz.rutaPPT, ComandosInterfaz.rutaExcel, b5)

            mensaje = 'Datos del PowerPoint exportados a Excel.'
            tkinter.messagebox.showinfo( "Completado" , mensaje )
        except Exception as e:
            mensaje = 'No fue posible completar la operación.'
            tkinter.messagebox.showwarning( "Error" , mensaje )
            
            tkinter.messagebox.showwarning( "Error" , e.__class__ )

            
    l1 = Label( tk, text='Herramienta para llenar archivos de Word y PowerPoint desde una plantilla de Excel.\n Desarrollado por Tilos lab para uso exclusivo de Energicol.' )
    l1.grid( row=0,column=0, columnspan=2 )

    b0 = Button( tk, text = 'Llenar Documentos', width = 60 )
    b0['command'] = lambda: switch( )
    b0.grid( row=1,column=0, columnspan=2 )

    # Frame

    frameCargue = LabelFrame( tk , text='Cargar Datos')
    frameCargue.grid( row=2,column=0,columnspan=2, rowspan = 3)

    #Titulos

    l1 = Label( frameCargue, text='Seleccionar un archivo Excel (.xlsx):' )
    l1.grid( row=0,column=0 )

    l2 = Label( frameCargue, text='Seleccionar un archivo Word (.docx):' )
    l2.grid( row=1,column=0 )

    l3 = Label( frameCargue, text='Seleccionar un archivo PowerPoint (.pptx):' )
    l3.grid( row=2,column=0 )

    # Botones

    b1 = Button( frameCargue, text = 'Seleccionar Un Archivo', width = 30 )
    b1['command'] = lambda: excel( b1)
    b1.grid( row=0,column=1 )

    b2 = Button( frameCargue, text = 'Seleccionar Un Archivo', width = 30 )
    b2['command'] = lambda: word( b2)
    b2.grid( row=1,column=1 )

    b3 = Button( frameCargue, text = 'Seleccionar Un Archivo', width = 30 )
    b3['command'] = lambda: ppt( b3)
    b3.grid( row=2,column=1 )

    #Fin de frame

    # Botones de Tareas
    b4 = Button( tk, text = 'Extraer los parámetros de word.', width = 30 )
    b4['state'] = 'disabled'
    b4['command'] = lambda: extraerWord( b4 )
    b4.grid( row=5,column=0 )

    b5 = Button( tk, text = 'Extraer los parámetros de PowerPoint.', width = 30 )
    b5['state'] = 'disabled'
    b5['command'] = lambda: extraerPpt( b5 )
    b5.grid( row=5,column=1 )

    b6 = Button( tk, text = 'Ubicar los parámetros en el word.', width = 30 )
    b6['state'] = 'disabled'
    b6['command'] = lambda: ubicarWord( b6)
    b6.grid( row=6,column=0 )

    b7 = Button( tk, text = 'Ubicar los parámetros en el PowerPoint.', width = 30 )
    b7['state'] = 'disabled'
    b7['command'] = lambda: ubicarPpt( b7 )
    b7.grid( row=6,column=1 )


    tk.mainloop()