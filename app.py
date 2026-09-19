import os
import tkinter as tk
from tkinter import filedialog, messagebox
# Importa as funções puras do arquivo processador.py
from processador import processar_pasta_relatorios, gerar_excel_fluencia

def executar_interface():
    pasta_selecionada = filedialog.askdirectory(title="Selecione a pasta com os relatórios do e-educa")
    if not pasta_selecionada:
        return

    dados_alunos, info_cabecalho, bimestres = processar_pasta_relatorios(pasta_selecionada)

    if not dados_alunos:
        messagebox.showwarning("Aviso", "Nenhum relatório HTML válido foi encontrado na pasta selecionada.")
        return

    caminho_saida = os.path.join(pasta_selecionada, "Relatorio_Anual_Fluencia.xlsx")
    gerar_excel_fluencia(dados_alunos, info_cabecalho, caminho_saida)

    bimestres_str = ", ".join(sorted(bimestres))
    messagebox.showinfo("Sucesso!", f"Relatório gerado com sucesso!\n\nBimestres identificados: {bimestres_str}\n\nArquivo salvo em:\n{caminho_saida}")

# Montagem da janela do Tkinter
root = tk.Tk()
root.title("Gerador de Relatórios e-educa")
root.geometry("450x220")
root.resizable(False, False)

label_titulo = tk.Label(root, text="Acompanhamento da Fluência Leitora", font=("Calibri", 14, "bold"))
label_titulo.pack(pady=15)

label_instrucao = tk.Label(root, text="Selecione a pasta onde salvou os relatórios HTML.", font=("Calibri", 10))
label_instrucao.pack(pady=5)

btn_gerar = tk.Button(root, text="📂 Selecionar Pasta e Gerar Excel", font=("Calibri", 11, "bold"), bg="#1E293B", fg="white", padx=15, pady=8, command=executar_interface)
btn_gerar.pack(pady=15)

root.mainloop()