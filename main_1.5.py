import customtkinter as ctk
from tkinter import filedialog
import pandas
from CTkMessagebox import CTkMessagebox
import webbrowser

FILENAME = ''
CAMPO_CORRENTE = 0

CSV_ANAGRAFICA = {
    'ana_cd': [],
    'ana_dn': [],
    'ana_parent_cd': [],
    'ana_tipo': [],
    'ana_ds': [],
    'ana_note': [],
}

COLONNE_ANACD = []  # ana_cd
COLONNE_ANADN = []  # ana_dn
COLONNE_ANAPARENT = []  # ana_parent_cd
NOME_ANAGRAFICA = ''  # ana_tipo
COLONNE_ANADS = []  # ana_ds
COLONNE_ANANOTE = []  # ana_note


def UploadAction(event=None):
    global FILENAME
    FILENAME = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry.configure(placeholder_text=FILENAME.split('/')[-1])


def open_webpage(event):
    url = "https://github.com/DavideFolk"
    webbrowser.open(url)


def aggiorna_interfaccia():
    global FILENAME
    global NOME_ANAGRAFICA

    if FILENAME == '':
        return

    if entry_anagrafica.get() == '':
        return

    NOME_ANAGRAFICA = entry_anagrafica.get()

    label.destroy()
    entry.destroy()
    buttonFile.destroy()
    button_procedi.destroy()
    entry_anagrafica.destroy()
    nome.destroy()

    app.geometry("800x300")

    excel_data_df = pandas.read_excel(FILENAME)

    nomi_colonne = excel_data_df.columns.tolist()

    # label
    label_anagrafica = ctk.CTkLabel(app, text="Crea l'anagrafica", font=('arial', 20))
    label_anagrafica.grid(row=0, column=0, padx=20, pady=(5, 10), columnspan=3)

    # menù
    option_menu = ctk.CTkOptionMenu(app, values=nomi_colonne, command=optionmenu_callback)
    option_menu.grid(row=2, column=0, padx=20, pady=10, sticky="ew", columnspan=2)

    # button prossimo campo
    button_campo = ctk.CTkButton(app, text='Prossimo campo', command=avanza_campo)
    button_campo.grid(row=2, column=2, padx=20, pady=10, sticky="ew")

    global label_testo_anagrafica
    label_testo_anagrafica = ctk.CTkLabel(app, text="ana_cd: ", font=('arial', 16), width=700, height=150, anchor='nw',
                                          justify='left')
    label_testo_anagrafica.grid(row=3, column=0, padx=20, pady=(5, 10), columnspan=3)

    # button crea anagrafica
    button_crea_anagr = ctk.CTkButton(app, text='Crea', command=crea_anagrafica)
    button_crea_anagr.grid(row=4, column=0, padx=20, sticky="ew", columnspan=2)

    # button resetta
    button_resetta = ctk.CTkButton(app, text='Resetta', command=resetta)
    button_resetta.grid(row=4, column=2, padx=20, sticky="ew")


def crea_anagrafica():
    global COLONNE_ANACD
    global COLONNE_ANADN
    global COLONNE_ANAPARENT
    global NOME_ANAGRAFICA
    global COLONNE_ANADS
    global COLONNE_ANANOTE

    if not COLONNE_ANADS:
        CTkMessagebox(message="Attenzione! compila i campi, ana_note può essere vuoto", icon="warning", option_1="Ok!",
                      title='Info', master=app)
        return

    excel_data_df = pandas.read_excel(FILENAME, dtype=str)

    for index, row in excel_data_df.iterrows():

        # ana_cd
        temp_list = []
        for colonna in COLONNE_ANACD:
            if not pandas.isna(row[colonna]):
                temp_list.append(str(row[colonna]).replace('\n', ' ').replace(';', ' ').strip())
            else:
                temp_list.append('')
        final_list = '|'.join(temp_list)
        CSV_ANAGRAFICA['ana_cd'].append(final_list)

        # ana_dn
        temp_list = []
        for colonna in COLONNE_ANADN:
            if not pandas.isna(row[colonna]):
                temp_list.append(str(row[colonna]).replace('\n', ' ').replace(';', ' ').strip())
            else:
                temp_list.append('')
        final_list = '|'.join(temp_list)
        CSV_ANAGRAFICA['ana_dn'].append(final_list)

        # ana_parent_cd
        temp_list = []
        for colonna in COLONNE_ANAPARENT:
            if not pandas.isna(row[colonna]):
                temp_list.append(str(row[colonna]).replace('\n', ' ').replace(';', ' ').strip())
            else:
                temp_list.append('')
        final_list = '|'.join(temp_list)
        CSV_ANAGRAFICA['ana_parent_cd'].append(final_list)

        # ana_tipo
        CSV_ANAGRAFICA['ana_tipo'].append(NOME_ANAGRAFICA)

        # ana_ds
        temp_list = []
        for colonna in COLONNE_ANADS:
            if not pandas.isna(row[colonna]):
                temp_list.append(str(row[colonna]).replace('\n', ' ').replace(';', ' ').strip())
            else:
                temp_list.append('')
        final_list = '|'.join(temp_list)
        CSV_ANAGRAFICA['ana_ds'].append(final_list)

        # ana_note
        temp_list = []
        for colonna in COLONNE_ANANOTE:
            if not pandas.isna(row[colonna]):
                temp_list.append(str(row[colonna]).replace('\n', ' ').replace(';', ' ').strip())
            else:
                temp_list.append('')
        final_list = '|'.join(temp_list)
        CSV_ANAGRAFICA['ana_note'].append(final_list)

    df = pandas.DataFrame(CSV_ANAGRAFICA)
    df.to_csv(NOME_ANAGRAFICA + '.csv', index=False, sep=';', header=False, mode='w')

    CTkMessagebox(message="Anagrafica creata!", icon="check", option_1="Ok!", title='Info', master=app)


def optionmenu_callback(value):
    dict_colonne = {
        0: COLONNE_ANACD,
        1: COLONNE_ANADN,
        2: COLONNE_ANAPARENT,
        3: COLONNE_ANADS,
        4: COLONNE_ANANOTE,
    }

    text = label_testo_anagrafica.cget("text")
    label_testo_anagrafica.configure(text=text + ' ' + value)
    dict_colonne[CAMPO_CORRENTE].append(value)


def resetta():
    global CAMPO_CORRENTE
    global COLONNE_ANADS
    global COLONNE_ANACD
    global COLONNE_ANADN
    global COLONNE_ANAPARENT
    global COLONNE_ANANOTE
    global CSV_ANAGRAFICA

    COLONNE_ANADS = []
    COLONNE_ANACD = []
    COLONNE_ANADN = []
    COLONNE_ANAPARENT = []
    COLONNE_ANANOTE = []

    CAMPO_CORRENTE = 0
    label_testo_anagrafica.configure(text="ana_cd:  ")

    CSV_ANAGRAFICA = {
        'ana_cd': [],
        'ana_dn': [],
        'ana_parent_cd': [],
        'ana_tipo': [],
        'ana_ds': [],
        'ana_note': [],
    }


def avanza_campo():
    global CAMPO_CORRENTE
    CAMPO_CORRENTE += 1

    text = label_testo_anagrafica.cget("text")

    match CAMPO_CORRENTE:
        case 1:
            label_testo_anagrafica.configure(text=text + '\nana_dn: ')
        case 2:
            label_testo_anagrafica.configure(text=text + '\nana_parent_cd: ')
        case 3:
            label_testo_anagrafica.configure(
                text=text + '\nana_tipo (nome anagrafica): ' + NOME_ANAGRAFICA + '\nana_ds: ')
        case 4:
            label_testo_anagrafica.configure(text=text + '\nana_note: ')


# ############################## customTkinter per l'interfaccia
app = ctk.CTk()
app.geometry("600x270")
app.title('Crea Anagrafica')

app.grid_columnconfigure((0, 1), weight=1)

# label
label = ctk.CTkLabel(app, text='Scegli il file excel', font=('arial', 24))
label.grid(row=0, column=0, padx=20, pady=(5, 30), columnspan=3)

# nome file
entry = ctk.CTkEntry(app, placeholder_text="path")
entry.grid(row=1, column=0, padx=20, pady=(10, 10), sticky="ew", columnspan=2)

# button file
buttonFile = ctk.CTkButton(app, text="File", command=UploadAction)
buttonFile.grid(row=1, column=2, padx=20, pady=(10, 10))

# nome anagrafica
entry_anagrafica = ctk.CTkEntry(app, placeholder_text="nome anagrafica")
entry_anagrafica.grid(row=2, column=0, padx=20, pady=(10, 10), sticky="ew", columnspan=2)

# button procedi
button_procedi = ctk.CTkButton(app, text='Procedi', command=aggiorna_interfaccia)
button_procedi.grid(row=3, column=0, padx=100, pady=(50, 0), sticky="ew", columnspan=3)

# nome
nome = ctk.CTkLabel(app, text='2024 - Davide Mazzone', font=('arial', 12))
nome.grid(row=4, column=2, padx=(25, 0), pady=(5, 30))
nome.bind("<Button-1>", open_webpage)

app.mainloop()

# pyinstaller --noconfirm --onefile --windowed --add-data "c:/users/work/desktop/code/python/CreaAnagrafiche/.venv/lib/site-packages/customtkinter;customtkinter/" "C:/Users/Work/Desktop/Code/Python/CreaAnagrafiche/main.py"
