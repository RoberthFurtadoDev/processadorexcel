# 🏠 Gerador de Cards de Registros de Imóveis

**Transforme planilhas Excel com registros imobiliários em cards textuais padronizados, prontos para impressão ou colagem.**  
Ideal para **padronizar descrições de imóveis** e **otimizar processos administrativos** com um formato limpo e automatizado.

---

## 📋 Índice
- [✨ Funcionalidades](#-funcionalidades)
- [⚙️ Pré-requisitos](#️-pré-requisitos)
- [🚀 Como Usar](#-como-usar)
- [🧾 Exemplo de Card Gerado](#-exemplo-de-card-gerado)
- [📥 Entrada Esperada](#-entrada-esperada)
- [🧠 Principais Funções e Lógica](#-principais-funções-e-lógica)
- [💡 Sugestões de Melhoria](#-sugestões-de-melhoria)
- [📄 Licença](#-licença)

---

## ✨ Funcionalidades

- 📂 **Leitura automática** de planilhas Excel (`.xlsx`) com dados imobiliários.  
- 🧱 **Geração de cards**: cada imóvel vira um bloco descritivo completo e amigável.  
- 🧮 **Padronização automática**:  
  - Conversão de áreas (m²)  
  - Pluralização correta de cômodos  
  - Formatação de números no padrão brasileiro  
- 📤 **Exportação formatada**:  
  - Cards em **coluna única**  
  - Altura de linha ajustada automaticamente  
  - Bordas finas e alinhamento central  

---

## ⚙️ Pré-requisitos

- Python **3.x**
- Biblioteca **openpyxl**

# Instale as dependências com:

```bash
pip install openpyxl
```

## 🚀 Como Usar

- Coloque sua planilha de imóveis (ex: imoveis.xlsx) na mesma pasta do script.

- Execute o comando:

- python processar.py [CAMINHO_DA_PLANILHA] [NOME_DA_ABA]


- Caso não forneça argumentos, o script usa os padrões:

- Planilha: Teste tit2.xlsx

Aba: Planilha1

## 📘 Exemplo:

- python processar.py "imoveis.xlsx" "Planilha1"


## 💾 O resultado será salvo automaticamente em:
- dados_formatados.xlsx

## 🧾 Exemplo de Card Gerado
- Residencial Sol. Apartamento, 73,76 m² de área total, 65 m² de área privativa, 2 quartos, 2 wcs.

- Andar alto; nascente.

- IPTU: 950,5
- Matrícula: 12345  Ofício: 2º.

## 📥 Entrada Esperada

- A planilha de entrada deve conter colunas (por índice fixo) correspondentes a:

- Empreendimento

- Tipo de imóvel

- Áreas (total, privativa, terreno)

- Quantidade e tipo de cômodos (quartos, wcs, varandas, salas, etc.)

- Descrição, IPTU, matrícula, ofício

- Consulte o dicionário IDX no código-fonte caso queira adaptar para outros layouts de planilha.

## 🧠 Principais Funções e Lógica

- Leitura otimizada com read_only=True e data_only=True (melhor performance).

- Conversão de cada linha (exceto cabeçalho) em texto descritivo formatado.

- Ajuste automático de altura e quebra de linha para não cortar conteúdo.

- Aplicação de bordas e centralização no Excel.

- Mensagens informativas de sucesso, erro e status da execução.

## 💡 Sugestões de Melhoria

- Parâmetros opcionais para layout (largura de coluna, fonte, aba de saída).

- Exportação para outros formatos: CSV, Markdown, PDF.

- Validação automática das colunas obrigatórias.

- Implementação de testes unitários nas funções de formatação.

## 📄 Licença

- Distribuído sob a Licença MIT.
- Sinta-se à vontade para colaborar, adaptar ou sugerir melhorias! 💻

## 🧑‍💻 Autor: Roberth Furtado
## 📅 Versão: 1.0
## 📦 Linguagem: Python 3
## 📚 Dependências: openpyxl



