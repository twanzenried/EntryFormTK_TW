import tkinter
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import os

# Funciones
def enter_data():

     accepted = accept_var.get()

     # if para que controle si los terminos han sido aceptados para utilizar el programa

     if accepted == "Has aceptado":
          # Info del usuario
          firstname = first_name_entry.get()
          lastname = last_name_entry.get()


          if firstname and lastname:
               title = title_combobox.get()
               age = age_spinbox.get()
               nationality = nationality_combobox.get()

               # Informacion del curso y semestre
               numcourses = numcourses_spinbox.get()
               numsemesters = numsemesters_spinbox.get()

               # Estado del registro
               registration_status = reg_status_var.get()

               print(f"Nombre: {firstname}; Apellido: {lastname}")
               print(f"Título: {title}; Edad: {age}; Nacionalidad: {nationality}")
               print(f"Cursos: {numcourses}; Semestres: {numsemesters}")
               print (f"Estado del registro: {registration_status}")
               print('-'* 40)

               filepath = "C:\\Users\\wanze\\OneDrive\\Escritorio\\pythonProject\\Python_Project" \
                          "\\data_entry_excels\\data_input.xlsx"

               if not os.path.exists(filepath):
                    workbook = openpyxl.Workbook()
                    sheet = workbook.active
                    heading = ["Nombre", "Apellido", "Título", "Edad", "Nacionalidad", "# de Cursos",
                               "# de Semestres", "Estado de registro"]
                    sheet.append(heading)
                    workbook.save(filepath)
               workbook = openpyxl.load_workbook(filepath)
               sheet = workbook.active
               sheet.append([firstname, lastname, title, age, nationality, numcourses, numsemesters,
                             registration_status])
               workbook.save(filepath)

          else:
               tkinter.messagebox.showwarning(title="Error",
                                              message="Nombre y Apellido no pueden tener campos vacios.")
     else:
          tkinter.messagebox.showwarning(title="Error",
                                         message="No has aceptado los términos y condiciones.")

# Crear ventana principal y el titulo
window = tkinter.Tk()
window.title("Formulario de Datos")

# Crear sub-ventana o frame dentro de la ventana
frame = tkinter.Frame(window)
frame.pack()

# Guardar info de usuario
user_info_frame = tkinter.LabelFrame(frame, text="Info de Usuario")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

# Etiqueta de nombre
first_name_label = tkinter.Label(user_info_frame, text="Nombre")
first_name_label.grid(row=0, column=0)

# Etiqueta de apellido
last_name_label = tkinter.Label(user_info_frame, text="Apellido")
last_name_label.grid(row=0, column=1)

# Crear entradas para que usuario introduzca informacion
first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

# Loop for para los padx y pady
for widget in user_info_frame.winfo_children():
     widget.grid_configure(padx=10, pady=5)

# Titulo
title_label = tkinter.Label(user_info_frame, text="Título")
title_combobox = ttk.Combobox(user_info_frame, values=["Sr.", "Sra.", "Dr.", ""])
title_label.grid(row=0, column=2)
title_combobox.grid(row=1, column=2)

# Edad
age_label = tkinter.Label(user_info_frame, text="Edad")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18, to=110)
age_label.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

# Nacionalidad
nationality_label = tkinter.Label(user_info_frame, text="Nacionalidad")
nationality_combobox = ttk.Combobox(user_info_frame, values=["Argentinx", "Chilenx", "Uruguayx", "Bolivianx", "Paraguayx"])
nationality_label.grid(row=2, column=1)
nationality_combobox.grid(row=3, column=1)

# Segunda label frame
courses_frame = tkinter.LabelFrame(frame)
courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

# Vincular checkbox con variable para poder aplicarle la funcion .get()
reg_status_var = tkinter.StringVar(value="No estás registrado")

# Etiqueta de registro
registered_label = tkinter.Label(courses_frame, text="Estado del Registro")
registered_check = tkinter.Checkbutton(courses_frame, text="Ya estas registrado", variable=reg_status_var,
                                       onvalue="Registrado", offvalue="Sin registrar")

registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

# Spinbox  numcourses
numcourses_label = tkinter.Label(courses_frame, text="# Formularios Completados")
numcourses_spinbox = tkinter.Spinbox(courses_frame, from_=0, to='infinity')
numcourses_label.grid(row=0, column=1)
numcourses_spinbox.grid(row=1, column=1)

# Numero de semestres (?)
numsemesters_label = tkinter.Label(courses_frame, text="# Semestres")
numsemesters_spinbox = tkinter.Spinbox(courses_frame, from_=0, to='infinity')
numsemesters_label.grid(row=0, column=2)
numsemesters_spinbox.grid(row=1, column=2)

# Loop for para acomodar el padding del frame inferior
for widget in courses_frame.winfo_children():
     widget.grid_configure(padx=2.9, pady=5)

# Aceptar los terminos de uso para permitir el acceso a la app
terms_frame = tkinter.LabelFrame(frame, text="Términos y Condiciones")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

# Variable de aceptacion de terminos para .get()
accept_var = tkinter.StringVar(value="No has aceptado")

# Boton para que el usuario acepte
terms_check = tkinter.Checkbutton(terms_frame, text="Yo acepto los térm. y condiciones", variable=accept_var,
                                  onvalue="Has aceptado", offvalue="No has aceptado")
terms_check.grid(row=0, column=0)

# Boton
button = tkinter.Button(frame, text="Inserte info.", command=enter_data)
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)

# Loop principal del programa
window.mainloop()