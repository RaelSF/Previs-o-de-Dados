import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from sklearn.linear_model import LinearRegression
from datetime import timedelta
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Função para habilitar o botão de carregamento
def habilitar_botao():
    nome_coluna_data = coluna_data_entry.get()
    nome_coluna_alvo = coluna_alvo_entry.get()

    if nome_coluna_data and nome_coluna_alvo:
        load_button.config(state="normal")
    else:
        load_button.config(state="disabled")

# Função para ajustar o modelo e fazer previsões
def ajustar_modelo_e_prever():
    nome_coluna_data = coluna_data_entry.get()
    nome_coluna_alvo = coluna_alvo_entry.get()

    if not nome_coluna_data or not nome_coluna_alvo:
        messagebox.showerror("Erro", "Por favor, forneça o nome da coluna da linha e da coluna de destino.")
        return

    file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if not file_path:
        return

    try:
        # Leitura do arquivo Excel
        df = pd.read_excel(file_path)

        # Conversão de datas para números inteiros representando dias desde a primeira data
        df[nome_coluna_data] = (df[nome_coluna_data] - df[nome_coluna_data].min()).dt.days

        # Separação das características (Data) e do alvo (Vendas)
        X = df[[nome_coluna_data]]
        y = df[nome_coluna_alvo]

        # Criação do modelo de regressão linear
        model = LinearRegression()
        model.fit(X, y)

        # Previsão para novos tempos (6, 7, 8 dias após a última data no conjunto de dados)
        ultima_data = df[nome_coluna_data].max()
        novos_tempos = [ultima_data + i for i in range(1, 4)]
        previsoes = model.predict([[tempo] for tempo in novos_tempos])

        # Exibição das previsões
        previsoes_text.delete(1.0, tk.END)
        for tempo, previsao in zip(novos_tempos, previsoes):
            previsoes_text.insert(tk.END, f"Previsão para o tempo {tempo} (em dias): {previsao:.2f}\n")

        # Gráfico aprimorado
        ax.clear()
        ax.scatter(X, y, label="Dados", color='blue')
        ax.scatter(novos_tempos, previsoes, label="Previsões", color='red', marker='x')
        ax.plot(X, model.predict(X), color='green', linestyle='--', label="Regressão Linear")
        ax.set_xlabel(f"{nome_coluna_data} (em dias desde a primeira data)")
        ax.set_ylabel(f"{nome_coluna_alvo}")
        ax.set_title("Previsão de Vendas com Regressão Linear")
        ax.legend()
        canvas.draw()

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao processar o arquivo Excel:\n{str(e)}")

# Criação da janela principal
root = tk.Tk()
root.title("Previsão de Vendas")

# Campos de entrada para o nome da coluna da linha e da coluna de destino
coluna_data_label = tk.Label(root, text="Nome da Coluna do tempo:")
coluna_data_label.pack()
coluna_data_entry = tk.Entry(root)
coluna_data_entry.pack()
coluna_data_entry.bind("<KeyRelease>", lambda event: habilitar_botao())

coluna_alvo_label = tk.Label(root, text="Nome da Coluna dos valores:")
coluna_alvo_label.pack()
coluna_alvo_entry = tk.Entry(root)
coluna_alvo_entry.pack()
coluna_alvo_entry.bind("<KeyRelease>", lambda event: habilitar_botao())

# Botão para carregar arquivo Excel (inicialmente desativado)
load_button = tk.Button(root, text="Carregar Arquivo Excel", command=ajustar_modelo_e_prever, state="disabled")
load_button.pack(pady=10)

# Área de texto para exibir previsões
previsoes_text = tk.Text(root, height=6, width=40)
previsoes_text.pack()

# Gráfico
fig, ax = plt.subplots(figsize=(6, 4))
canvas = FigureCanvasTkAgg(fig, master=root)
canvas_widget = canvas.get_tk_widget()
canvas_widget.pack()

root.mainloop()