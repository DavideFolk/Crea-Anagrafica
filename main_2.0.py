import customtkinter as ctk
from tkinter import filedialog
import pandas
from CTkMessagebox import CTkMessagebox
import webbrowser
from PIL import Image
import json
import sys
import os

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # quando gira come exe
    except Exception:
        base_path = os.path.abspath(".") # quando gira in pycharm

    return os.path.join(base_path, relative_path)

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

dict_colonne = {
    0: COLONNE_ANACD,
    1: COLONNE_ANADN,
    2: COLONNE_ANAPARENT,
    3: COLONNE_ANADS,
    4: COLONNE_ANANOTE,
}

icona_esporta = ctk.CTkImage(
    light_image=Image.open(resource_path("icons/export.png")),
    size=(25, 25)
)

icona_importa = ctk.CTkImage(
    light_image=Image.open(resource_path("icons/import.png")),
    size=(25, 25)
)

icona_crea = ctk.CTkImage(
    light_image=Image.open(resource_path("icons/crea.png")),
    size=(20, 20)
)

icona_indietro = ctk.CTkImage(
    light_image=Image.open(resource_path("icons/undo.png")),
    size=(20, 20)
)

icona_reset = ctk.CTkImage(
    light_image=Image.open(resource_path("icons/reset.png")),
    size=(20, 20)
)


def UploadAction(event=None):
    global FILENAME
    FILENAME = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry.configure(state="normal")  # per permettere modifiche temporanee via codice
    entry.delete(0, "end")
    entry.configure(placeholder_text=FILENAME.split('/')[-1])
    entry.configure(state="readonly")


def open_webpage(event):
    url = "https://github.com/DavideFolk"
    webbrowser.open(url)


def aggiorna_interfaccia():
    global NOME_ANAGRAFICA

    if FILENAME == '':
        return

    if entry_anagrafica.get() == '':
        return

    NOME_ANAGRAFICA = entry_anagrafica.get()

    # Distruggi i widget precedenti
    for widget in app.winfo_children():
        widget.destroy()

    app.geometry("900x550")

    excel_data_df = pandas.read_excel(FILENAME)
    nomi_colonne = excel_data_df.columns.tolist()

    # Header con titolo e pulsanti
    frame_header = ctk.CTkFrame(app, fg_color="transparent")
    frame_header.grid(
        row=0,
        column=0,
        padx=30,
        pady=(15, 8),  # ↓ meno spazio sotto
        sticky="ew",
        columnspan=4
    )

    label_anagrafica = ctk.CTkLabel(
        frame_header,
        text="Seleziona i campi",
        font=('Segoe UI', 24, 'bold')
    )
    label_anagrafica.pack(side="left")

    # Frame per i pulsanti importa/esporta
    frame_actions = ctk.CTkFrame(frame_header, fg_color="transparent")
    frame_actions.pack(side="right")

    button_importa = ctk.CTkButton(
        frame_actions,
        text="Importa",
        image=icona_importa,
        width=110,
        height=36,
        corner_radius=8,
        compound="left",
        font=('Segoe UI', 13),
        command=importa
    )
    button_importa.pack(side="left", padx=5)

    button_esporta = ctk.CTkButton(
        frame_actions,
        text="Esporta",
        image=icona_esporta,
        width=110,
        height=36,
        corner_radius=8,
        compound="left",
        font=('Segoe UI', 13),
        command=esporta
    )
    button_esporta.pack(side="left", padx=5)

    # Card per la selezione del campo
    frame_selection = ctk.CTkFrame(app, corner_radius=12, fg_color=("gray90", "gray20"))
    frame_selection.grid(
        row=1,
        column=0,
        padx=30,
        pady=(0, 8),
        sticky="ew",
        columnspan=4
    )

    label_campo = ctk.CTkLabel(
        frame_selection,
        text="Campo da mappare",
        font=('Segoe UI', 14),
        text_color=("gray40", "gray60")
    )
    label_campo.grid(
        row=0,
        column=0,
        padx=20,
        pady=(10, 4),
        sticky="w"
    )

    option_menu = ctk.CTkOptionMenu(
        frame_selection,
        values=nomi_colonne,
        command=optionmenu_callback,
        height=40,
        width=420,
        corner_radius=8,
        font=('Segoe UI', 14),
        dropdown_font=('Segoe UI', 13)
    )
    option_menu.grid(
        row=1,
        column=0,
        padx=20,
        pady=(0, 10),
        sticky="ew"
    )

    button_campo = ctk.CTkButton(
        frame_selection,
        text='Prossimo campo →',
        command=avanza_campo,
        height=40,
        corner_radius=8,
        font=('Segoe UI', 14, 'bold'),
        fg_color=("#3b8ed0", "#1f6aa5")
    )
    button_campo.grid(
        row=1,
        column=1,
        padx=(10, 20),
        pady=(0, 10),
        sticky="ew"
    )

    frame_selection.grid_columnconfigure(0, weight=3)
    frame_selection.grid_columnconfigure(1, weight=1)

    # Card per il testo dell'anagrafica
    global label_testo_anagrafica
    frame_anagrafica = ctk.CTkFrame(app, corner_radius=12, fg_color=("gray90", "gray20"))
    frame_anagrafica.grid(
        row=2,
        column=0,
        padx=30,
        pady=(0, 8),
        sticky="nsew",
        columnspan=4
    )

    label_header_ana = ctk.CTkLabel(
        frame_anagrafica,
        text="Anteprima Anagrafica",
        font=('Segoe UI', 14, 'bold'),
        text_color=("gray30", "gray70")
    )
    label_header_ana.pack(
        anchor='w',
        padx=20,
        pady=(10, 4)
    )

    label_testo_anagrafica = ctk.CTkLabel(
        frame_anagrafica,
        text="ana_cd: ",
        font=('Consolas', 15),
        anchor='nw',
        justify='left',
        wraplength=800,
        text_color=("gray20", "gray90")
    )
    label_testo_anagrafica.pack(fill="both", expand=True, padx=20, pady=(0, 15))

    # Imposta il peso delle righe per il layout responsive
    app.grid_rowconfigure(0, weight=0)
    app.grid_rowconfigure(1, weight=0)
    app.grid_rowconfigure(2, weight=1)  # solo anteprima cresce
    app.grid_rowconfigure(3, weight=0)

    # Frame per i pulsanti inferiori
    frame_buttons = ctk.CTkFrame(app, fg_color="transparent")
    frame_buttons.grid(
        row=3,
        column=0,
        padx=30,
        pady=(0, 15),
        sticky="ew",
        columnspan=4
    )

    button_resetta = ctk.CTkButton(
        frame_buttons,
        text="Resetta",
        image=icona_reset,
        compound="left",
        height=44,
        corner_radius=8,
        font=('Segoe UI', 14),
        fg_color=("gray70", "gray30"),
        hover_color=("gray60", "gray40"),
        command=resetta
    )
    button_resetta.pack(side="left", fill="x", expand=True, padx=(0, 10))

    button_annulla = ctk.CTkButton(
        frame_buttons,
        text="Annulla",
        image=icona_indietro,
        compound="left",
        height=44,
        corner_radius=8,
        font=('Segoe UI', 14),
        fg_color=("gray70", "gray30"),
        hover_color=("gray60", "gray40"),
        command=annulla
    )
    button_annulla.pack(side="left", fill="x", expand=True, padx=5)

    button_crea_anagr = ctk.CTkButton(
        frame_buttons,
        text="Crea Anagrafica",
        image=icona_crea,
        compound="left",
        height=44,
        corner_radius=8,
        font=('Segoe UI', 14, 'bold'),
        fg_color=("#2ecc71", "#27ae60"),
        hover_color=("#27ae60", "#229954"),
        command=crea_anagrafica
    )
    button_crea_anagr.pack(side="left", fill="x", expand=True, padx=(10, 0))

    # Configura le colonne per il layout responsive
    app.grid_columnconfigure(0, weight=1)


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
                temp_list.append(str(row[colonna]).replace('\n', ' ').replace(';', ' ').replace('|', ' ').replace('"', '').strip())
            else:
                temp_list.append('')
        final_list = '|'.join(temp_list)
        CSV_ANAGRAFICA['ana_cd'].append(final_list)

        # ana_dn
        temp_list = []
        for colonna in COLONNE_ANADN:
            if not pandas.isna(row[colonna]):
                temp_list.append(str(row[colonna]).replace('\n', ' ').replace(';', ' ').replace('|', ' ').replace('"', '').strip())
            else:
                temp_list.append('')
        final_list = '|'.join(temp_list)
        CSV_ANAGRAFICA['ana_dn'].append(final_list)

        # ana_parent_cd
        temp_list = []
        for colonna in COLONNE_ANAPARENT:
            if not pandas.isna(row[colonna]):
                temp_list.append(str(row[colonna]).replace('\n', ' ').replace(';', ' ').replace('|', ' ').replace('"', '').strip())
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
                temp_list.append(str(row[colonna]).replace('\n', ' ').replace(';', ' ').replace('|', ' ').replace('"', '').strip())
            else:
                temp_list.append('')
        final_list = '|'.join(temp_list)
        CSV_ANAGRAFICA['ana_ds'].append(final_list)

        # ana_note
        temp_list = []
        for colonna in COLONNE_ANANOTE:
            if not pandas.isna(row[colonna]):
                temp_list.append(str(row[colonna]).replace('\n', ' ').replace(';', ' ').replace('|', ' ').replace('"', '').strip())
            else:
                temp_list.append('')
        final_list = '|'.join(temp_list)
        CSV_ANAGRAFICA['ana_note'].append(final_list)

    df = pandas.DataFrame(CSV_ANAGRAFICA)
    df.to_csv(NOME_ANAGRAFICA + '.csv', index=False, sep=';', header=False, mode='w')

    CTkMessagebox(message="Anagrafica creata!", icon="check", option_1="Ok!", title='Info', master=app)


def optionmenu_callback(value):
    dict_colonne[CAMPO_CORRENTE].append(value)
    aggiorna_label()


def resetta():
    global CAMPO_CORRENTE
    global CSV_ANAGRAFICA

    COLONNE_ANADS.clear()
    COLONNE_ANACD.clear()
    COLONNE_ANADN.clear()
    COLONNE_ANAPARENT.clear()
    COLONNE_ANANOTE.clear()

    CAMPO_CORRENTE = 0

    CSV_ANAGRAFICA = {
        'ana_cd': [],
        'ana_dn': [],
        'ana_parent_cd': [],
        'ana_tipo': [],
        'ana_ds': [],
        'ana_note': [],
    }
    aggiorna_label()


def annulla():
    valori = dict_colonne.get(CAMPO_CORRENTE, [])

    if not valori:
        CTkMessagebox(
            message="I campi sono già vuoti!",
            icon="warning",
            option_1="Ok!",
            title="Info",
            master=app
        )
        return

    valori.pop()
    aggiorna_label()


def aggiorna_label():
    CAMPI = [
        ("ana_cd", COLONNE_ANACD, 0),
        ("ana_dn", COLONNE_ANADN, 1),
        ("ana_parent_cd", COLONNE_ANAPARENT, 2),
        ("ana_ds", COLONNE_ANADS, 3),
        ("ana_note", COLONNE_ANANOTE, 4),
    ]

    righe = []

    for nome, valori, indice in CAMPI:
        # mostra sempre se ha valori
        if valori:
            righe.append(f"{nome}: " + " | ".join(valori))
        # mostra SOLO se è il campo corrente
        elif indice == CAMPO_CORRENTE:
            righe.append(f"{nome}: ")

    # campo anagrafica
    if CAMPO_CORRENTE > 2:
        righe.insert(
            3,
            f"ana_tipo (nome anagrafica): {NOME_ANAGRAFICA}"
        )

    label_testo_anagrafica.configure(text="\n".join(righe))


def avanza_campo():
    global CAMPO_CORRENTE

    if CAMPO_CORRENTE < 4:
        CAMPO_CORRENTE += 1
        aggiorna_label()

def esporta():
    stato = {
        "COLONNE_ANACD": COLONNE_ANACD,
        "COLONNE_ANADN": COLONNE_ANADN,
        "COLONNE_ANAPARENT": COLONNE_ANAPARENT,
        "NOME_ANAGRAFICA": NOME_ANAGRAFICA,
        "COLONNE_ANADS": COLONNE_ANADS,
        "COLONNE_ANANOTE": COLONNE_ANANOTE,
    }

    with open("stato_anagrafica.json", "w", encoding="utf-8") as f:
        json.dump(stato, f, indent=2, ensure_ascii=False)

    CTkMessagebox(message="Export eseguito!", icon="check", option_1="Ok!", title='Info', master=app)

def importa():
    global NOME_ANAGRAFICA

    with open("stato_anagrafica.json", "r", encoding="utf-8") as f:
        stato = json.load(f)

    COLONNE_ANACD.clear()
    COLONNE_ANACD.extend(stato.get("COLONNE_ANACD", []))

    COLONNE_ANADN.clear()
    COLONNE_ANADN.extend(stato.get("COLONNE_ANADN", []))

    COLONNE_ANAPARENT.clear()
    COLONNE_ANAPARENT.extend(stato.get("COLONNE_ANAPARENT", []))

    COLONNE_ANADS.clear()
    COLONNE_ANADS.extend(stato.get("COLONNE_ANADS", []))

    COLONNE_ANANOTE.clear()
    COLONNE_ANANOTE.extend(stato.get("COLONNE_ANANOTE", []))

    NOME_ANAGRAFICA = stato.get("NOME_ANAGRAFICA", "")

    aggiorna_label()


# ############################## customTkinter per l'interfaccia
app = ctk.CTk()
app.geometry("700x550")
app.title('Crea Anagrafica')

app.grid_columnconfigure(0, weight=1)
app.grid_rowconfigure(1, weight=1)

# Header
frame_header = ctk.CTkFrame(app, fg_color="transparent")
frame_header.grid(row=0, column=0, padx=40, pady=(10, 6), sticky="ew")

label = ctk.CTkLabel(
    frame_header,
    text='Crea Anagrafica',
    font=('Segoe UI', 32, 'bold')
)
label.pack(anchor="w")

label_subtitle = ctk.CTkLabel(
    frame_header,
    text='Carica un file Excel per iniziare',
    font=('Segoe UI', 14),
    text_color=("gray40", "gray60")
)
label_subtitle.pack(anchor="w", pady=(5, 0))

# Card principale
frame_main = ctk.CTkFrame(app, corner_radius=16, fg_color=("gray90", "gray20"))
frame_main.grid(row=1, column=0, padx=40, pady=(0, 20), sticky="nsew")

# Sezione file
label_file = ctk.CTkLabel(
    frame_main,
    text="File Excel",
    font=('Segoe UI', 14, 'bold'),
    anchor="w"
)
label_file.grid(row=0, column=0, padx=30, pady=(30, 10), sticky="w", columnspan=2)

frame_file = ctk.CTkFrame(frame_main, fg_color="transparent")
frame_file.grid(row=1, column=0, padx=30, pady=(0, 25), sticky="ew", columnspan=2)

entry = ctk.CTkEntry(
    frame_file,
    placeholder_text="Nessun file selezionato",
    height=44,
    corner_radius=8,
    font=('Segoe UI', 13),
    state="readonly"
)
entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

buttonFile = ctk.CTkButton(
    frame_file,
    text="Sfoglia",
    command=UploadAction,
    height=44,
    width=120,
    corner_radius=8,
    font=('Segoe UI', 13, 'bold')
)
buttonFile.pack(side="left")

# Divisore visuale
frame_divider = ctk.CTkFrame(frame_main, height=1, fg_color=("gray70", "gray40"))
frame_divider.grid(row=2, column=0, padx=30, pady=(0, 25), sticky="ew", columnspan=2)

# Sezione nome anagrafica
label_anagrafica = ctk.CTkLabel(
    frame_main,
    text="Nome Anagrafica",
    font=('Segoe UI', 14, 'bold'),
    anchor="w"
)
label_anagrafica.grid(row=3, column=0, padx=30, pady=(0, 10), sticky="w", columnspan=2)

entry_anagrafica = ctk.CTkEntry(
    frame_main,
    placeholder_text="Inserisci il nome dell'anagrafica",
    height=44,
    corner_radius=8,
    font=('Segoe UI', 13)
)
entry_anagrafica.grid(row=4, column=0, padx=30, pady=(0, 30), sticky="ew", columnspan=2)

# Configura colonne del frame_main
frame_main.grid_columnconfigure(0, weight=1)

# Button procedi
button_procedi = ctk.CTkButton(
    app,
    text='Procedi →',
    command=aggiorna_interfaccia,
    height=50,
    corner_radius=10,
    font=('Segoe UI', 16, 'bold'),
    fg_color=("#2ecc71", "#27ae60"),
    hover_color=("#27ae60", "#229954"),
)
button_procedi.grid(row=2, column=0, padx=40, pady=(0, 20), sticky="ew")

# Footer
frame_footer = ctk.CTkFrame(app, fg_color="transparent")
frame_footer.grid(row=3, column=0, padx=40, pady=(0, 20), sticky="ew")

nome = ctk.CTkLabel(
    frame_footer,
    text='© 2025 - Davide Mazzone',
    font=('Segoe UI', 11),
    text_color=("gray50", "gray60"),
    cursor="hand2"
)
nome.pack(side="right")
nome.bind("<Button-1>", open_webpage)

app.mainloop()

# pyinstaller --noconfirm --onefile --windowed --add-data "icons;icons" main_2.0.py
