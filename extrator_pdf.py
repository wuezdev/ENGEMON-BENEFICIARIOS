import tabula
import pandas as pd
import os
import PyPDF2

class ExtratorPDF:
    def __init__(self, caminho_pdf):
        self.caminho_pdf = caminho_pdf
        self.dados = None
        
    def extrair_tabelas(self):
        """Extrai tabelas do PDF usando tabula-py"""
        try:
            # Extrair todas as tabelas do PDF
            tabelas = tabula.read_pdf(
                self.caminho_pdf,
                pages='all',
                multiple_tables=True,
                lattice=True  # Para tabelas com linhas delimitadoras
            )
            
            # Verificar se alguma tabela foi extraída
            if not tabelas:
                raise Exception("Nenhuma tabela encontrada no PDF.")
            
            # Processar as tabelas encontradas
            df_completo = None
            
            for i, tabela in enumerate(tabelas):
                # Limpar nomes das colunas (remover espaços extras)
                tabela.columns = [col.strip() if isinstance(col, str) else col for col in tabela.columns]
                
                # Se for a primeira tabela, usar como base
                if i == 0:
                    df_completo = tabela.copy()
                # Caso contrário, anexar à base
                else:
                    df_completo = pd.concat([df_completo, tabela], ignore_index=True)
            
            # Filtrar colunas desejadas, se estiverem presentes
            colunas_desejadas = [
                'Status', 'Tipo', 'Nome', 'CPF', 'Nascimento', 
                'Titular', 'Parentesco', 'Matricula', 'Data de inclusão'
            ]
            
            # Verificar quais colunas existem no DataFrame
            colunas_existentes = [col for col in colunas_desejadas if col in df_completo.columns]
            
            # Se não encontrou nenhuma coluna esperada, tente usar auto-detecção
            if not colunas_existentes:
                # Tentar novamente sem usar lattice para casos onde a tabela não tem linhas delimitadoras
                tabelas = tabula.read_pdf(
                    self.caminho_pdf,
                    pages='all',
                    multiple_tables=True,
                    lattice=False   
                )
                
                df_completo = pd.concat(tabelas, ignore_index=True) if tabelas else None
                
                if df_completo is None:
                    raise Exception("Falha na extração de tabelas do PDF.")
                
                # Verificar novamente as colunas existentes
                colunas_existentes = [col for col in colunas_desejadas if col in df_completo.columns]
            
            # Filtrar apenas as colunas desejadas que existem no DataFrame
            if colunas_existentes:
                self.dados = df_completo[colunas_existentes].copy()
            else:
                # Se ainda não encontrou as colunas, mantém todas as colunas
                self.dados = df_completo.copy()
                
            return True
        
        except Exception as e:
            print(f"Erro ao extrair tabelas do PDF: {str(e)}")
            return False
            
    def obter_dados(self):
        """Retorna os dados extraídos"""
        return self.dados