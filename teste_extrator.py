from extrator_pdf import ExtratorPDF

# Substitua pelo caminho do seu arquivo PDF de exemplo
caminho_pdf = "beneficiarios.pdf"

extrator = ExtratorPDF(caminho_pdf)
sucesso = extrator.extrair_tabelas()

if sucesso:
    dados = extrator.obter_dados()
    print("Extração concluída com sucesso!")
    print(f"Colunas encontradas: {list(dados.columns)}")
    print(f"Número de registros: {len(dados)}")
    print("\nPrimeiras 5 linhas:")
    print(dados.head())
else:
    print("Falha na extração do PDF.")