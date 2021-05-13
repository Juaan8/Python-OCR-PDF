import tabula
import pandas as pd
from tkinter import *
from tkinter import filedialog 



def abrePDF():
	global file

	file=filedialog.askopenfilename(title="Abrir", filetypes=(("Ficheros PDF", "*.pdf"), ("Todos los ficheros", "*.*")))
	(miRuta.set(file))


def guardarArchivo():
	global archivo_guardado

	archivo_guardado=filedialog.asksaveasfilename(title="Guardar en", defaultextension=".xlsx", 
		filetypes=(("Ficheros xlsx", "*.xlsx"), ("Ficheros csv", "*.csv"), ("Todos los ficheros", "*.*")))
	(miArchivo.set(archivo_guardado))
	

def codigoBoton():

	try:

		file
		answer.config(text="")

	except NameError:

		answer.config(text="¡ SELECCIONE UN PDF !", fg="red")

		return False


	try:

		int(pagina.get())	
		answer.config(text="")

	except ValueError:

		answer.config(text="¡ DEBE INGRESAR EL NUMERO DE PAGINA !", fg="red")

		return False


	try:
		archivo_guardado
		answer.config(text="")

	except NameError: 

		answer.config(text="¡ SELECCIONE DONDE GUARDAR !", fg="red")

		return False




	
	tables = tabula.read_pdf(file, pages =int(pagina.get()), multiple_tables = True, stream=True)


	tabula.convert_into(file, "reporte.csv", pages = int(pagina.get()))

	df = pd.read_csv('reporte.csv')

	writer = pd.ExcelWriter(archivo_guardado)

	df.to_excel(writer, index=False)

	writer.save() 

	answer.config(text="¡ ARCHIVO CONVERTIDO !", fg="green")



root=Tk()
root.title("OCR")


miFrame=Frame(root, width=300, height=150)
miFrame.pack()


miFrame2=Frame(root)
miFrame2.pack()

miImagen=PhotoImage(file="imagenes/logo.png")
Label(miFrame, image=miImagen).place(x=-15, y=0)


miFrame3=Frame(root)
miFrame3.pack()

miRuta=StringVar()

botonAbrirPDF = Button(miFrame3, text="Abrir PDF", command=abrePDF)
botonAbrirPDF.grid(row=0, column=1, sticky="n", padx=10, pady=10)

paginaLabel=Label(miFrame3, textvariable=miRuta)
paginaLabel.grid(row=1, column=1, sticky="e", padx=10, pady=10)



miFrame4=Frame(root)
miFrame4.pack()

pagina=StringVar()


paginaLabel=Label(miFrame4, text="Pagina:")
paginaLabel.grid(row=2, column=0, sticky="e", padx=10, pady=10)

cuadroPaginas=Entry(miFrame4, textvariable=pagina)
cuadroPaginas.grid(row=2, column=1, padx=10, pady=10)


miFrame5=Frame(root)
miFrame5.pack()

answer = Label(miFrame5, text='')
answer.grid(row=3, column=1, padx=10, pady=10)



miFrame6=Frame(root)
miFrame6.pack()

miArchivo=StringVar()

botonGuardar=Button(miFrame6, text="Guardar en...", command=guardarArchivo)
botonGuardar.grid(row=4, column=1, sticky="s", padx=10, pady=10)

paginaLabel=Label(miFrame6, textvariable=miArchivo)
paginaLabel.grid(row=5, column=1, sticky="e", padx=10, pady=10)

botonConvertir=Button(miFrame6, text="Convertir", command=codigoBoton)
botonConvertir.grid(row=6, column=1, sticky="s", padx=10, pady=10)


root.mainloop()