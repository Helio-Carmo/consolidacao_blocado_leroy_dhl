import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import webbrowser
from datetime import datetime

# === FUNÇÃO PARA LER CAMINHOS SALVOS ===
def carregar_caminhos_salvos():
    try:
        with open("caminhos.txt", "r") as f:
            linhas = f.read().splitlines()
            return linhas if len(linhas) == 4 else ["", "", "", ""]
    except:
        return ["", "", "", ""]

# === FUNÇÃO PARA SALVAR CAMINHOS ===
def salvar_caminhos(caminhos):
    with open("caminhos.txt", "w") as f:
        f.write("\n".join(caminhos))

# === FUNÇÃO PRINCIPAL DE CONSOLIDAÇÃO ===
def consolidar_blocado(caminhos, label_resultado, label_stats):
    try:
        # === LÊ AS BASES ===
        if not caminhos[0].endswith('.xlsx') or not caminhos[1].endswith('.xlsx') or not caminhos[2].endswith('.xlsx'):
            raise Exception("Selecione arquivos Excel (.xlsx) válidos para as três primeiras entradas.")

        base_blocado = pd.read_excel(caminhos[0])
        tipologia = pd.read_excel(caminhos[1])
        empilhamento = pd.read_excel(caminhos[2])

        # Padroniza colunas
        base_blocado.columns = base_blocado.columns.str.strip().str.lower()
        empilhamento.columns = empilhamento.columns.str.strip().str.upper()
        tipologia.columns = tipologia.columns.str.strip()

        colunas_esperadas = {'produto': 'Produto', 'lote': 'Lote', 'posição no depósito': 'posição no depósito'}
        for original, novo in colunas_esperadas.items():
            if original in base_blocado.columns:
                base_blocado.rename(columns={original: novo}, inplace=True)
            elif novo not in base_blocado.columns:
                raise Exception(f"Coluna '{novo}' não encontrada na base blocado")

        if 'UMA' not in empilhamento.columns:
            raise Exception("Coluna 'UMA' não encontrada na base de empilhamento")

        empilhamento_pal = empilhamento[empilhamento["UMA"] == "PAL"].copy()
        empilhamento_pal.rename(columns={"MATERIAL": "Produto", "PALLET - EMPILHAMENTO MÁXIMO": "Empilhamento"}, inplace=True)

        blocado_emp = pd.merge(base_blocado, empilhamento_pal[["Produto", "Empilhamento"]], on="Produto", how="left")
        blocado_emp = pd.merge(blocado_emp, tipologia[["Pos.depós.", "TP"]], left_on="posição no depósito", right_on="Pos.depós.", how="left")

        blocado_emp["Capacidade_base"] = blocado_emp["TP"].str.extract(r'B(\d+)').astype(float)
        blocado_emp["Empilhamento"] = blocado_emp["Empilhamento"].fillna(1)
        blocado_emp["Capacidade_total_pallets"] = blocado_emp["Empilhamento"] * blocado_emp["Capacidade_base"]

        ocupacao_resumo = blocado_emp.groupby(["posição no depósito", "Produto", "Lote"]).agg({
            "Produto": "count",
            "Capacidade_total_pallets": "first",
            "TP": "first",
            "Empilhamento": "first"
        }).rename(columns={"Produto": "Pallets_ocupados"}).reset_index()

        produto_multiplas_posicoes = ocupacao_resumo.groupby(["Produto", "Lote"]).filter(lambda x: len(x) > 1)

        def simular_consolidacao(df_produto):
            df = df_produto.set_index("posição no depósito").copy()
            df["Capacidade_restante"] = df["Capacidade_total_pallets"] - df["Pallets_ocupados"]

            movimentacoes = []
            liberadas = []
            ja_movimentados = set()

            while True:
                origem = df[df["Pallets_ocupados"] > 0].sort_values("Pallets_ocupados").head(1)
                destino = df[df["Capacidade_restante"] > 0].sort_values("Pallets_ocupados", ascending=False).head(1)

                if origem.empty or destino.empty:
                    break

                origem_idx = origem.index[0]
                destino_idx = destino.index[0]

                if origem_idx == destino_idx or (origem_idx, destino_idx) in ja_movimentados:
                    break
                if df.at[origem_idx, "Lote"] != df.at[destino_idx, "Lote"]:
                    break

                mover = min(origem["Pallets_ocupados"].values[0], destino["Capacidade_restante"].values[0])
                if mover <= 0:
                    break

                df.at[origem_idx, "Pallets_ocupados"] -= mover
                df.at[destino_idx, "Pallets_ocupados"] += mover
                df["Capacidade_restante"] = df["Capacidade_total_pallets"] - df["Pallets_ocupados"]

                ocupacao_percentual = df.at[destino_idx, "Pallets_ocupados"] / df.at[destino_idx, "Capacidade_total_pallets"]

                movimentacoes.append({
                    "Produto": df.at[origem_idx, "Produto"],
                    "Empilhamento": df.at[origem_idx, "Empilhamento"],
                    "Lote": df.at[origem_idx, "Lote"],
                    "Posição Origem": origem_idx,
                    "Tipologia Origem": df.at[origem_idx, "TP"],
                    "Posição Destino": destino_idx,
                    "Tipologia Destino": df.at[destino_idx, "TP"],
                    "Pallets Movidos": mover,
                    "Capacidade Total Destino": df.at[destino_idx, "Capacidade_total_pallets"],
                    "Ocupação Posição Destino": f"{ocupacao_percentual:.0%}"
                })

                ja_movimentados.add((origem_idx, destino_idx))

            for pos, row in df.iterrows():
                if row["Pallets_ocupados"] == 0:
                    liberadas.append(pos)

            return movimentacoes, liberadas

        sugestoes = []
        posicoes_liberadas = []

        for (_, _), grupo in produto_multiplas_posicoes.groupby(["Produto", "Lote"]):
            consolidado, liberadas = simular_consolidacao(grupo)
            sugestoes.extend(consolidado)
            posicoes_liberadas.extend(liberadas)

        df_sugestoes = pd.DataFrame(sugestoes)
        df_liberadas = pd.DataFrame({"Posições_liberadas": posicoes_liberadas})

        if df_sugestoes.empty:
            messagebox.showinfo("Resultado", "Nenhuma sugestão de consolidação foi gerada.")
            return

        pasta_destino = caminhos[3] if caminhos[3] else str(Path.home() / "Desktop")
        datahora = datetime.now().strftime("%Y-%m-%d_%Hh%Mm")
        nome_arquivo = os.path.join(pasta_destino, f"consolidacao_{datahora}.xlsx")

        with pd.ExcelWriter(nome_arquivo, engine="xlsxwriter") as writer:
            df_sugestoes.to_excel(writer, sheet_name="Sugestoes", index=False)
            df_liberadas.to_excel(writer, sheet_name="Posicoes Liberadas", index=False)

            resumo_df = pd.DataFrame([{
                "Total Pallets Movimentados": df_sugestoes["Pallets Movidos"].sum(),
                "Total Posições Envolvidas": df_sugestoes[["Posição Origem", "Posição Destino"]].nunique().sum(),
                "Total Posições Liberadas": len(df_liberadas)
            }])
            resumo_df.to_excel(writer, sheet_name="Resumo", index=False)

            # Adiciona tabela de posições liberadas por tipologia
            if not df_liberadas.empty:
                df_liberadas = df_liberadas.merge(tipologia[["Pos.depós.", "TP"]], left_on="Posições_liberadas", right_on="Pos.depós.", how="left")
                resumo_tipologia = df_liberadas["TP"].value_counts().reset_index()
                resumo_tipologia.columns = ["TIPOLOGIA", "Total Posições Liberadas"]
                resumo_tipologia = resumo_tipologia.sort_values("TIPOLOGIA")

                resumo_tipologia.to_excel(writer, sheet_name="Resumo", startrow=4, index=False)

            workbook = writer.book
            center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
            writer.sheets["Sugestoes"].set_column("A:Z", 20, center_format)
            writer.sheets["Posicoes Liberadas"].set_column("A:A", 25, center_format)

        label_resultado.config(fg="green", text=f"Arquivo gerado: {nome_arquivo}", cursor="hand2")
        label_resultado.bind("<Button-1>", lambda e: webbrowser.open(nome_arquivo))

        total_movimentacoes = df_sugestoes["Pallets Movidos"].sum()
        posicoes_env = df_sugestoes[["Posição Origem", "Posição Destino"]].nunique().sum()
        total_liberadas = len(df_liberadas)

        label_stats.config(fg="black", text=f"Pallets movimentados: {total_movimentacoes} | Posições envolvidas: {posicoes_env} | Posições liberadas: {total_liberadas}")

    except Exception as e:
        label_resultado.config(fg="red", text=f"Erro: {str(e)}")

# === CRIAR INTERFACE TKINTER ===
root = tk.Tk()
root.resizable(False, False)  # Impede redimensionamento horizontal e vertical
root.attributes("-toolwindow", 1)  # Remove botão de maximizar (em Windows)
root.title("Consolidação Blocado")

# Iniciar a janela no centro da tela
largura = 700
altura = 320

largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()
x = (largura_tela - largura) // 2
y = (altura_tela - altura) // 2
root.geometry(f"{largura}x{altura}+{x}+{y}")
#====================================
root.configure(bg="#e3e3e3")

try:
    root.iconbitmap("icon.ico")
except Exception as e:
    print(f"Erro ao carregar ícone: {e}")

caminhos = carregar_caminhos_salvos()
labels = ["Base Blocado", "Base Tipologia", "Base Empilhamento", "Salvar em"]
entradas = []

for i, label in enumerate(labels):
    tk.Label(root, text=label, bg="#e3e3e3", font=("Arial", 10, "bold")).place(x=50, y=30 + i*50)
    entrada = tk.Entry(root, width=60)
    entrada.place(x=200, y=30 + i*50)
    entrada.insert(0, caminhos[i])
    entradas.append(entrada)

    def selecionar_arquivo(e=i):
        if e < 3:
            caminho = filedialog.askopenfilename()
        else:
            caminho = filedialog.askdirectory()
        if caminho:
            entradas[e].delete(0, tk.END)
            entradas[e].insert(0, caminho)

    tk.Button(root, text="Selecionar", command=selecionar_arquivo).place(x=560, y=30 + i*50)

label_resultado = tk.Label(root, text="", bg="#e3e3e3", font=("Arial", 10, "bold"))
label_resultado.place(x=50, y=260)

label_stats = tk.Label(root, text="", bg="#e3e3e3", font=("Arial", 10))
label_stats.place(x=50, y=290)

def ao_clicar_gerar():
    novos_caminhos = [entrada.get() for entrada in entradas]
    salvar_caminhos(novos_caminhos)
    consolidar_blocado(novos_caminhos, label_resultado, label_stats)

tk.Button(root, text="Gerar Arquivo", command=ao_clicar_gerar, bg="#002a5e", fg="white", font=("Arial", 10, "bold"), width=20).place(x=270, y=220)

# Botão de Ajuda
def mostrar_ajuda():
    messagebox.showinfo("Ajuda", "Selecione as três bases conforme o formato padrão:\n"
                         "- Base Blocado: colunas Produto, Lote, posição no depósito\n"
                         "- Tipologia: colunas Pos.depós., TP\n"
                         "- Empilhamento: colunas UMA, MATERIAL, PALLET - EMPILHAMENTO MÁXIMO\n"
                         "\nClique em 'Gerar Arquivo' para consolidar posições com o mesmo produto e lote.")

tk.Button(root, text="Ajuda", command=mostrar_ajuda).place(x=650, y=10)

root.mainloop()
