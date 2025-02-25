# Extrator de Beneficiários - Engemon

Aplicação para extração automática de dados de tabelas de beneficiários a partir de arquivos PDF e exportação para formato Excel.

## Funcionalidades

- Interface gráfica intuitiva
- Seleção de arquivos PDF
- Extração automática de tabelas
- Exportação para Excel
- Visualização de progresso durante o processamento

## Requisitos

- Python 3.7+
- Bibliotecas Python (instaladas automaticamente):
  - PyPDF2
  - pandas
  - tabula-py
  - openpyxl
  - PyQt5

## Instalação

1. Clone o repositório:
```
git clone https://github.com/seuusuario/ENGEMON-BENEFICIARIOS.git
cd ENGEMON-BENEFICIARIOS
```

2. Instale as dependências:
```
pip install -r requirements.txt
```

## Uso

Execute o programa principal:
```
python main.py
```

### Instruções de uso:
1. Clique em "Selecionar Arquivo PDF"
2. Escolha o arquivo PDF contendo a tabela de beneficiários
3. Clique em "Processar"
4. Aguarde o processamento
5. O arquivo Excel será gerado automaticamente

## Estrutura do Projeto

```
engemon_beneficiarios/
├── main.py               # Ponto de entrada do programa
├── extrator_pdf.py       # Módulo para extração de dados do PDF
├── conversor_excel.py    # Módulo para conversão para Excel
├── gui.py                # Interface gráfica
└── requirements.txt      # Dependências do projeto
```

## Notas de Desenvolvimento

- A extração de tabelas utiliza a biblioteca `tabula-py`, que requer Java 8+ instalado no sistema
- A interface gráfica foi desenvolvida com PyQt5 para compatibilidade multiplataforma