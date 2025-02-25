import sys
import os
import pandas as pd
import tabula
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                             QHBoxLayout, QFileDialog, QLabel, QWidget, QProgressBar)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

class ExtratorThread(QThread):
    """Thread para processar a extração sem congelar a interface."""
    progresso = pyqtSignal(int)
    concluido = pyqtSignal(str)
    erro = pyqtSignal(str)
    
    def __init__(self, arquivo_pdf):
        super().__init__()
        self.arquivo_pdf = arquivo_pdf
        
    def run(self):
        try:
            # Extrair as tabelas do PDF
            self.progresso.emit(20)
            tabelas = tabula.read_pdf(self.arquivo_pdf, pages='all', multiple_tables=True)
            
            self.progresso.emit(50)
            
            # Identificar a tabela principal com os beneficiários
            tabela_beneficiarios = None
            for tabela in tabelas:
                # Verificamos se a tabela contém as colunas esperadas
                colunas = list(tabela.columns)
                if 'Status' in colunas and 'Nome' in colunas and 'CPF' in colunas:
                    tabela_beneficiarios = tabela
                    break
            
            self.progresso.emit(70)
            
            if tabela_beneficiarios is None:
                self.erro.emit("Não foi possível encontrar a tabela de beneficiários no PDF.")
                return
                
            # Remover a coluna "Data de bloqueio" se existir
            if 'Data de bloqueio' in tabela_beneficiarios.columns:
                tabela_beneficiarios = tabela_beneficiarios.drop(columns=['Data de bloqueio'])
                
            self.progresso.emit(90)
            
            # Salvar como Excel
            nome_arquivo_base = os.path.splitext(os.path.basename(self.arquivo_pdf))[0]
            caminho_saida = os.path.join(os.path.dirname(self.arquivo_pdf), f"{nome_arquivo_base}_extraido.xlsx")
            tabela_beneficiarios.to_excel(caminho_saida, index=False)
            
            self.progresso.emit(100)
            self.concluido.emit(caminho_saida)
            
        except Exception as e:
            self.erro.emit(f"Erro durante a extração: {str(e)}")


class BeneficiariosApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Extrator de Beneficiários - Engemon")
        self.setGeometry(100, 100, 600, 300)
        self.arquivo_pdf = None
        
        # Widget principal
        widget_central = QWidget()
        self.setCentralWidget(widget_central)
        
        # Layout principal
        layout_principal = QVBoxLayout(widget_central)
        
        # Área de seleção de arquivo
        layout_arquivo = QHBoxLayout()
        
        self.label_arquivo = QLabel("Nenhum arquivo selecionado")
        self.label_arquivo.setStyleSheet("padding: 8px; background-color: #f0f0f0; border-radius: 4px;")
        
        self.botao_selecionar = QPushButton("Selecionar PDF")
        self.botao_selecionar.setStyleSheet("padding: 8px; background-color: #4CAF50; color: white;")
        self.botao_selecionar.clicked.connect(self.selecionar_arquivo)
        
        layout_arquivo.addWidget(self.label_arquivo, 3)
        layout_arquivo.addWidget(self.botao_selecionar, 1)
        
        # Barra de progresso
        self.barra_progresso = QProgressBar()
        self.barra_progresso.setVisible(False)
        
        # Área de status
        self.label_status = QLabel("Aguardando arquivo PDF...")
        self.label_status.setAlignment(Qt.AlignCenter)
        
        # Botão de processamento
        self.botao_processar = QPushButton("Extrair para Excel")
        self.botao_processar.setStyleSheet("padding: 10px; background-color: #2196F3; color: white; font-weight: bold;")
        self.botao_processar.clicked.connect(self.processar_arquivo)
        self.botao_processar.setEnabled(False)
        
        # Adicionar elementos ao layout principal
        layout_principal.addLayout(layout_arquivo)
        layout_principal.addWidget(self.barra_progresso)
        layout_principal.addWidget(self.label_status)
        layout_principal.addWidget(self.botao_processar)
        
        # Ajustar margens e espaçamento
        layout_principal.setContentsMargins(20, 20, 20, 20)
        layout_principal.setSpacing(15)
        
    def selecionar_arquivo(self):
        arquivo, _ = QFileDialog.getOpenFileName(self, "Selecionar arquivo PDF", "", "Arquivos PDF (*.pdf)")
        
        if arquivo:
            self.arquivo_pdf = arquivo
            self.label_arquivo.setText(os.path.basename(arquivo))
            self.botao_processar.setEnabled(True)
            self.label_status.setText("Arquivo selecionado. Clique em 'Extrair para Excel' para continuar.")
    
    def processar_arquivo(self):
        if not self.arquivo_pdf:
            self.label_status.setText("Por favor, selecione um arquivo PDF primeiro.")
            return
        
        # Desabilitar botão durante o processamento
        self.botao_processar.setEnabled(False)
        self.botao_selecionar.setEnabled(False)
        self.barra_progresso.setVisible(True)
        self.barra_progresso.setValue(0)
        self.label_status.setText("Processando...")
        
        # Iniciar thread de extração
        self.thread_extracao = ExtratorThread(self.arquivo_pdf)
        self.thread_extracao.progresso.connect(self.atualizar_progresso)
        self.thread_extracao.concluido.connect(self.extracao_concluida)
        self.thread_extracao.erro.connect(self.extracao_erro)
        self.thread_extracao.start()
    
    def atualizar_progresso(self, valor):
        self.barra_progresso.setValue(valor)
    
    def extracao_concluida(self, caminho_saida):
        self.label_status.setText(f"Extração concluída com sucesso! Arquivo salvo em:\n{caminho_saida}")
        self.botao_processar.setEnabled(True)
        self.botao_selecionar.setEnabled(True)
        
        # Opção para abrir o diretório
        diretorio = os.path.dirname(caminho_saida)
        self.botao_processar.setText("Extrair novamente")
        
        # Criar botão para abrir o diretório
        botao_abrir = QPushButton("Abrir pasta com arquivo")
        botao_abrir.setStyleSheet("padding: 10px; background-color: #FF9800; color: white;")
        botao_abrir.clicked.connect(lambda: os.startfile(diretorio) if os.name == 'nt' else os.system(f'xdg-open "{diretorio}"'))
        
        # Adicionar botão ao layout
        layout = self.centralWidget().layout()
        layout.addWidget(botao_abrir)
    
    def extracao_erro(self, mensagem):
        self.label_status.setText(f"Erro: {mensagem}")
        self.botao_processar.setEnabled(True)
        self.botao_selecionar.setEnabled(True)
        self.barra_progresso.setVisible(False)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = BeneficiariosApp()
    janela.show()
    sys.exit(app.exec_())