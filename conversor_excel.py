import pandas as pd
import os
from datetime import datetime

class ConversorExcel:
    def __init__(self, dados, diretorio_saida=None):
        self.dados = dados
        
        # Se não for especificado um diretório, usa o diretório atual
        if diretorio_saida is None:
            self.diretorio_saida = os.getcwd()
        else:
            self.diretorio_saida = diretorio_saida
            
            # Criar o diretório se não existir
            if not os.path.exists(diretorio_saida):
                os.makedirs(diretorio_saida)
    
    def converter(self, nome_arquivo=None):
        """Converte os dados para Excel"""
        try:
            # Se não foi especificado um nome, cria um nome com data e hora
            if nome_arquivo is None:
                data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_arquivo = f"beneficiarios_{data_hora}.xlsx"
            
            # Verificar se o nome do arquivo termina com .xlsx
            if not nome_arquivo.endswith('.xlsx'):
                nome_arquivo += '.xlsx'
                
            # Caminho completo para o arquivo
            caminho_arquivo = os.path.join(self.diretorio_saida, nome_arquivo)
            
            # Salvar DataFrame como Excel
            self.dados.to_excel(caminho_arquivo, index=False)
            
            return caminho_arquivo
        
        except Exception as e:
            print(f"Erro ao converter para Excel: {str(e)}")
            return None