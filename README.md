# 📦 Consolidação Blocado - Leroy/DHL

Este projeto tem como objetivo **otimizar a ocupação de posições de blocado em um depósito**, consolidando pallets do mesmo produto e lote para liberar espaço. A aplicação foi desenvolvida em **Python com interface Tkinter**.

---

## 💡 Funcionalidades

- Leitura de três planilhas:
  - **Base Blocado**
  - **Base Tipologia**
  - **Base Empilhamento**
- Identifica produtos e lotes que estão ocupando múltiplas posições.
- Simula movimentações para consolidação.
- Gera arquivo `.xlsx` com:
  - Sugestões de movimentação
  - Lista de posições liberadas
  - Resumo geral e por Tipologia

---

## 🛠️ Como usar

1. Execute o programa (`main.py` ou o `.exe` gerado).
2. Selecione os arquivos necessários:
   - **Base Blocado:** deve conter as colunas `Produto`, `Lote`, `posição no depósito`.
   - **Base Tipologia:** deve conter `Pos.depós.` e `TP`.
   - **Base Empilhamento:** deve conter `UMA`, `MATERIAL`, e `PALLET - EMPILHAMENTO MÁXIMO`.
3. Escolha onde salvar o resultado.
4. Clique em **"Gerar Arquivo"**.
5. O programa criará um arquivo Excel com os resultados.

---

## 📁 Estrutura do Projeto


---

## 🧰 Tecnologias usadas

- Python 3.x
- Tkinter (GUI)
- Pandas
- XlsxWriter

---

## 🤝 Agradecimentos

Este projeto foi desenvolvido com apoio do **ChatGPT (OpenAI)**, que auxiliou em diversas etapas da lógica de consolidação, construção da interface gráfica, tratamento de dados com Pandas e empacotamento para `.exe`. 💡🚀

---

## 📝 Observações

- O programa salva os caminhos usados anteriormente no arquivo `caminhos.txt`.
- Ícone personalizado incluído na interface.
- Janela centralizada e com tamanho ajustado automaticamente.
- Exportação em formato `.xlsx`.

---

## 💬 Suporte

Em caso de dúvidas ou melhorias, sinta-se à vontade para abrir uma [issue](https://github.com/Helio-Carmo/consolidacao_blocado_leroy_dhl/issues) ou contribuir com um Pull Request.

---

## 👨‍💻 Autor

**Helio Carmo**  
[GitHub](https://github.com/Helio-Carmo)
