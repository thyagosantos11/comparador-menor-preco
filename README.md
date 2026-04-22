# 📊 Comparador de Menor Preço

Aplicação desktop para análise automática de cotações de fornecedores, desenvolvida para o **Mercado do Pai**.

Carrega uma planilha Excel com preços de múltiplos fornecedores e gera automaticamente um relatório formatado destacando o menor preço por produto.

---

## ✨ Funcionalidades

- Carrega planilhas `.xlsx` no padrão **RELAÇÃO DE COMPRA**
- Detecta fornecedores automaticamente pelas colunas
- Calcula o menor preço por produto
- Gera relatório Excel formatado com destaque visual (verde) para o menor preço
- Exibe ranking de fornecedores por número de menores preços
- Interface desktop intuitiva — sem necessidade de instalar Python

---

## 🖥️ Como usar (cliente final)

1. Baixe o executável em [Releases](../../releases)
2. Clique duas vezes no `Comparador de Precos.exe`
3. Selecione o arquivo `RELAÇÃO DE COMPRA.xlsx`
4. Clique em **Gerar Relatório Comparativo**
5. O arquivo `resultado_menor_preco.xlsx` será salvo na pasta escolhida

---

## 🛠️ Desenvolvimento

### Pré-requisitos

- Python 3.10+
- pip

### Instalação

```bash
git clone https://github.com/seu-usuario/comparador-menor-preco.git
cd comparador-menor-preco
pip install -r requirements.txt
```

### Rodar em modo desenvolvimento

```bash
python app.py
```

### Gerar o executável `.exe`

```bash
pyinstaller --onefile --windowed --name "Comparador de Precos" app.py
```

O `.exe` será gerado em `dist/`.

---

## 📁 Estrutura do projeto

```
comparador-menor-preco/
├── app.py                  # Interface desktop (CustomTkinter)
├── requirements.txt        # Dependências Python
├── README.md               # Este arquivo
├── .gitignore              # Arquivos ignorados pelo Git
└── exemplo/
    └── RELAÇÃO DE COMPRA (exemplo).xlsx   # Planilha de exemplo
```

---

## 📋 Formato esperado da planilha

| Venda | Produto | Quant | Fornecedor A | Fornecedor B | ... | Obs |
|-------|---------|-------|--------------|--------------|-----|-----|
| ...   | ...     | ...   | 2.50         | 3.10         | ... | ... |

- A primeira coluna deve ser **Venda**
- A segunda coluna deve ser **Produto**
- A terceira coluna deve ser **Quant**
- As colunas do meio são os **fornecedores** (detectados automaticamente)
- A última coluna deve ser **Obs**

Valores inválidos aceitos: `f`, `F`, `?` ou célula vazia (tratados como "sem cotação").

---

## 🧰 Tecnologias

- [Python 3](https://python.org)
- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) — interface desktop
- [pandas](https://pandas.pydata.org) — leitura e processamento de dados
- [openpyxl](https://openpyxl.readthedocs.io) — geração do Excel formatado
- [PyInstaller](https://pyinstaller.org) — empacotamento em `.exe`

---

## 📄 Licença

Projeto privado — todos os direitos reservados.
