from extrator_pdf import ExtratorPDF
from conversor_excel import ConversorExcel

# Substitua pelo caminho do seu arquivo PDF de exemplo
caminho_pdf = "beneficiarios.pdf"

# Extrair dados do PDF
extrator = ExtratorPDF(caminho_pdf)
if extrator.extrair_tabelas():
    dados = extrator.obter_dados()
    
    # Converter para Excel
    conversor = ConversorExcel(dados)
    caminho_excel = conversor.converter()
    
    if caminho_excel:
        print(f"Arquivo Excel gerado com sucesso: {caminho_excel}")
    else:
        print("Falha ao gerar arquivo Excel.")
else:
    print("Falha na extração do PDF.")