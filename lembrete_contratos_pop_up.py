import pandas as pd
import tkinter as tk
from tkinter import PhotoImage
import schedule
import time
import threading
import logging
from datetime import datetime, timedelta
from PIL import Image, ImageTk

EXCEL_PATH = "contratos.xlsx"
LOGO_PATH = "ARPE.jpg"

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Cache do DataFrame
contratos_df = None

def carregar_contratos():
    global contratos_df
    if contratos_df is None:
        try:
            contratos_df = pd.read_excel(EXCEL_PATH)
            logging.info("Contratos carregados com sucesso.")
        except Exception as e:
            logging.error(f"Erro ao carregar contratos: {e}")
            contratos_df = pd.DataFrame()
    return contratos_df

def salvar_contratos():
    try:
        contratos_df.to_excel(EXCEL_PATH, index=False)
        logging.info("Alterações salvas com sucesso.")
    except Exception as e:
        logging.error(f"Erro ao salvar contratos: {e}")

def verificar_contratos():
    df = carregar_contratos()
    hoje = datetime.now().date()

    for idx, row in df.iterrows():
        status = str(row['Status']).strip().lower()
        if status != 'resolvido':
            try:
                data_termino = pd.to_datetime(row['Vigencia_Termino']).date()
                if hoje >= data_termino - timedelta(days=90):
                    exibir_popup(idx, row)
            except Exception as e:
                logging.warning(f"Erro ao processar linha {idx}: {e}")

def exibir_popup(idx, row):
    def marcar_resolvido():
        contratos_df.at[idx, 'Status'] = 'Resolvido'
        salvar_contratos()
        root.destroy()

    def marcar_em_andamento():
        contratos_df.at[idx, 'Status'] = 'Em andamento'
        salvar_contratos()
        root.destroy()

    root = tk.Tk()
    root.title("⚠️ Alerta de Contrato - ARPE")
    root.configure(bg="white")

    try:
        logo_img = Image.open(LOGO_PATH)
        logo_img = logo_img.resize((120, 120))
        logo = ImageTk.PhotoImage(logo_img)
        logo_label = tk.Label(root, image=logo, bg="white")
        logo_label.image = logo
        logo_label.pack(pady=(10, 0))
    except Exception as e:
        logging.warning(f"Erro ao carregar logo: {e}")

    msg = (
        f"Empresa: {row['Empresa']}\n"
        f"Contrato: {row['Contrato']}\n"
        f"Vigência até: {row['Vigencia_Termino']}\n\n"
        f"Status atual: {row['Status']}"
    )

    tk.Label(root, text=msg.strip(), bg="white", fg="#333",
             font=("Arial", 11), justify="left", padx=20, pady=10).pack()

    btn_frame = tk.Frame(root, bg="white")
    btn_frame.pack(pady=10)

    tk.Button(btn_frame, text="Em andamento", bg="orange", fg="white",
              font=("Arial", 10), command=marcar_em_andamento,
              padx=12, pady=6).pack(side="left", padx=10)
    tk.Button(btn_frame, text="OK - Resolvido", bg="green", fg="white",
              font=("Arial", 10), command=marcar_resolvido,
              padx=12, pady=6).pack(side="right", padx=10)

    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    root.mainloop()

def thread_scheduler():
    schedule.every(1).hours.do(verificar_contratos)
    verificar_contratos()

    while True:
        schedule.run_pending()
        time.sleep(1)

# Inicia verificação em paralelo
threading.Thread(target=thread_scheduler, daemon=True).start()

# Evita que o script termine imediatamente
while True:
    time.sleep(60)
