# Processador de Dados de Imóveis

Sistema de automação para extração e formatação de dados de imóveis a partir de planilhas Excel, gerando saída padronizada com células mescladas e centralizadas.

## 📋 Descrição

Este projeto foi desenvolvido para atender uma demanda específica da empresa, automatizando o processamento manual de dados de imóveis. A solução reduz significativamente o tempo operacional e elimina perdas financeiras decorrentes de erros humanos no tratamento de grandes volumes de dados.

## 🚀 Tecnologias

- **Python 3.13+**
- **openpyxl** - Manipulação de arquivos Excel (.xlsx)
- **Arquitetura Hexagonal** - Separação em camadas (domínio, aplicação, infraestrutura)

## 📁 Estrutura do Projeto

processar_refatorado/
├── domain/
│ ├── init.py
│ ├── entities.py # Entidades de domínio (Imovel)
│ └── services.py # Serviços (formatação de números e rótulos)
├── application/
│ ├── init.py
│ └── use_cases.py # Caso de uso: ProcessarImoveisUseCase
├── infrastructure/
│ ├── init.py
│ ├── excel_reader.py # Adaptador de leitura Excel
│ └── excel_writer.py # Adaptador de escrita Excel
├── config.py # Configurações centralizadas
├── main.py # Ponto de entrada
├── .gitignore
├── README.md
└── requirements.txt

text

## ⚙️ Instalação

### Pré-requisitos
- Python 3.13 ou superior
- pip

### Passos

1. Clone o repositório:
git clone https://github.com/seu-usuario/processar-imoveis.git
cd processar-imoveis

text

2. Crie um ambiente virtual (recomendado):
python -m venv venv
source venv/bin/activate # Linux/Mac
venv\Scripts\activate # Windows

text

3. Instale as dependências:
pip install -r requirements.txt

text

## 🎯 Uso

### Execução básica
python main.py "Teste tit2.xlsx" "Planilha1"

text

### Parâmetros
- **Argumento 1** (opcional): Caminho do arquivo Excel de entrada
  - Padrão: `Teste tit2.xlsx`
- **Argumento 2** (opcional): Nome da planilha a processar
  - Padrão: `Planilha1`

### Variáveis de ambiente
Crie um arquivo `.env` (já ignorado no `.gitignore`) para sobrescrever configurações:

INPUT_FILE=seu_arquivo.xlsx
INPUT_SHEET=NomeDaPlanilha
OUTPUT_FILE=saida_customizada.xlsx

text

### Exemplo de saída
O script gera um arquivo `dados_formatados.xlsx` com:
- Uma única planilha "Registros"
- Cada imóvel em uma célula da coluna A (A1, A2, A3, ...)
- Texto centralizado com quebra de linha automática
- Bordas ao redor de cada "card"
- Formatação brasileira: vírgula como separador decimal, m² para áreas

## 📊 Formato de Entrada

A planilha de entrada deve seguir este layout (índices 0-based):

| Coluna | Índice | Descrição |
|--------|--------|-----------|
| A | 0 | Empreendimento |
| B | 1 | Tipo do imóvel |
| D | 3 | Área total |
| G | 6 | Área privativa |
| J | 9 | Área do terreno |
| L | 11 | Quantidade de quartos |
| M | 12 | Rótulo "quartos" |
| N | 13 | Vaga (texto livre) |
| ... | ... | ... |
| Y | 24 | Descrição adicional |
| AA | 26 | IPTU |
| AC | 28 | Matrícula |
| AE | 30 | Ofício |

## 🧪 Testes

(Futuro: adicionar testes unitários com pytest)

pytest tests/

text

## 🏗️ Arquitetura

O projeto segue princípios da **Arquitetura Hexagonal**:

- **Domínio** (`domain/`): Regras de negócio puras, sem dependências externas
- **Aplicação** (`application/`): Orquestração dos casos de uso
- **Infraestrutura** (`infrastructure/`): Adaptadores para Excel (I/O)
- **Config** (`config.py`): Configurações centralizadas e injetáveis

### Benefícios
- ✅ Testabilidade: camadas isoladas facilitam mocks
- ✅ Manutenibilidade: mudanças localizadas
- ✅ Escalabilidade: fácil adicionar novos adaptadores (JSON, CSV, DB)

## 🤝 Contribuindo

1. Fork o projeto
2. Crie uma branch para sua feature (`git checkout -b feature/MinhaFeature`)
3. Commit suas mudanças (`git commit -m 'feat: adiciona nova feature'`)
4. Push para a branch (`git push origin feature/MinhaFeature`)
5. Abra um Pull Request

### Convenção de commits
- `feat:` nova funcionalidade
- `fix:` correção de bug
- `docs:` apenas documentação
- `refactor:` refatoração sem mudar comportamento
- `test:` adiciona ou corrige testes

## 📝 Licença

Este projeto é proprietário e de uso interno da empresa.

## 👤 Autor

**Seu Nome**
- GitHub: [@seu-usuario](https://github.com/seu-usuario)
- LinkedIn: [Seu Perfil](https://linkedin.com/in/seu-perfil)

## 🙏 Agradecimentos

Desenvolvido para otimizar processos internos e reduzir perdas operacionais.
requirements.txt
Liste as dependências do projeto.​

text
openpyxl==3.1.2
Comandos para subir no Git
bash
git init
git add .gitignore README.md requirements.txt
git add domain/ application/ infrastructure/ config.py main.py
git commit -m "feat: adiciona processador de imóveis com arquitetura hexagonal"
git branch -M main
git remote add origin https://github.com/seu-usuario/processar-imoveis.git
git push -u origin main