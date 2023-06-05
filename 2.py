import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QGridLayout, QLabel, QLineEdit, \
    QComboBox, QSpinBox, QDoubleSpinBox, QPushButton, QGroupBox, QTableWidget, QTableWidgetItem

RESISTIVIDADE_COBRE = 0.00000172
RESISTIVIDADE_ALUMINIO = 0.00000282

class CalculoQuedaTensão(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Cálculo de Queda de Tensão')
        self.setGeometry(100, 100, 800, 600)

        self.linha_atual = 0

        self.grupo_entrada = self.criar_grupo_entrada()
        self.grupo_historico = self.criar_grupo_historico()

        self.central_widget = QWidget()
        self.central_layout = QVBoxLayout()
        self.central_layout.addWidget(self.grupo_entrada)
        self.central_layout.addWidget(self.grupo_historico)
        self.central_widget.setLayout(self.central_layout)
        self.setCentralWidget(self.central_widget)

    def calcular_queda_tensão(self):
        nome = self.nome_edit.text()
        ligacao = self.esquema_ligação_combo.currentText()
        tensao_alimentacao = self.tensão_m_edit.value()
        fp = self.fator_potência_edit.value()
        potencia_ativa = self.potência_ativa_edit.value()
        comprimento_condutor = self.comprimento_condutor_edit.value()
        bitola = float(self.seção_condutor_combo.currentText())
        cabo = self.material_cabo_combo.currentText()

        self.atualizar_tabela_historico(nome, ligacao, tensao_alimentacao, fp, potencia_ativa,
                                        comprimento_condutor, bitola, cabo)

    def atualizar_tabela_historico(self, nome, ligacao, tensao_alimentacao, fp, potencia_ativa,
                                   comprimento_condutor, bitola, cabo):
        self.tabela_historico.setRowCount(self.linha_atual + 1)

        queda_100 = 0.0
        queda_v = 0.0

        # Cálculo da corrente
        corrente = self.calcular_corrente(potencia_ativa, fp, tensao_alimentacao)

        # Cálculo da queda de tensão e tensão a jusante
        if cabo == 'Cobre':
            queda_100 = (corrente * comprimento_condutor * RESISTIVIDADE_COBRE * 100) / bitola
        elif cabo == 'Alumínio':
            queda_100 = (corrente * comprimento_condutor * RESISTIVIDADE_ALUMINIO * 100) / bitola
        queda_v = tensao_alimentacao - (tensao_alimentacao * queda_100 / 100)
        queda_100 = round(queda_100, 2)
        queda_v = round(queda_v, 2)

        # Atualização da tabela de histórico
        self.tabela_historico.setItem(self.linha_atual, 0, QTableWidgetItem(nome))
        self.tabela_historico.setItem(self.linha_atual, 1, QTableWidgetItem(ligacao))
        self.tabela_historico.setItem(self.linha_atual, 2, QTableWidgetItem(str(tensao_alimentacao)))
        self.tabela_historico.setItem(self.linha_atual, 3, QTableWidgetItem(str(fp)))
        self.tabela_historico.setItem(self.linha_atual, 4, QTableWidgetItem(str(potencia_ativa)))
        self.tabela_historico.setItem(self.linha_atual, 5, QTableWidgetItem(str(comprimento_condutor)))
        self.tabela_historico.setItem(self.linha_atual, 6, QTableWidgetItem(str(bitola)))
        self.tabela_historico.setItem(self.linha_atual, 7, QTableWidgetItem(cabo))
        self.tabela_historico.setItem(self.linha_atual, 8, QTableWidgetItem(str(queda_100)))
        self.tabela_historico.setItem(self.linha_atual, 9, QTableWidgetItem(str(queda_v)))

        self.linha_atual += 1

    def calcular_corrente(self, potencia_ativa, fp, tensao_alimentacao):
        corrente = potencia_ativa / (fp * tensao_alimentacao)
        return round(corrente, 2)

    def criar_grupo_historico(self):
        grupo_historico = QGroupBox('Histórico')
        layout_historico = QVBoxLayout()

        self.tabela_historico = QTableWidget()
        self.tabela_historico.setColumnCount(10)
        self.tabela_historico.setHorizontalHeaderLabels(
            ['Nome', 'Esquema de Ligação', 'Tensão de Alimentação (V)', 'Fator de Potência', 'Potência Ativa (W)',
             'Comprimento do Condutor (m)', 'Seção do Condutor (mm²)', 'Material do Cabo', 'Queda de Tensão (%)',
             'Tensão a Jusante (V)'])
        self.tabela_historico.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabela_historico.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabela_historico.setColumnWidth(0, 100)
        self.tabela_historico.setColumnWidth(1, 150)
        self.tabela_historico.setColumnWidth(2, 200)
        self.tabela_historico.setColumnWidth(3, 150)
        self.tabela_historico.setColumnWidth(4, 150)
        self.tabela_historico.setColumnWidth(5, 200)
        self.tabela_historico.setColumnWidth(6, 200)
        self.tabela_historico.setColumnWidth(7, 150)
        self.tabela_historico.setColumnWidth(8, 150)
        self.tabela_historico.setColumnWidth(9, 150)

        layout_historico.addWidget(self.tabela_historico)
        grupo_historico.setLayout(layout_historico)

        return grupo_historico

    def criar_grupo_entrada(self):
        self.nome_edit = QLineEdit()
        self.esquema_ligação_combo = QComboBox()
        self.esquema_ligação_combo.addItems(['Estrela', 'Triângulo'])
        self.tensão_m_edit = QDoubleSpinBox()
        self.tensão_m_edit.setMinimum(0.1)
        self.tensão_m_edit.setMaximum(9999.9)
        self.tensão_m_edit.setDecimals(1)
        self.fator_potência_edit = QDoubleSpinBox()
        self.fator_potência_edit.setMinimum(0.1)
        self.fator_potência_edit.setMaximum(1.0)
        self.fator_potência_edit.setSingleStep(0.1)
        self.potência_ativa_edit = QDoubleSpinBox()
        self.potência_ativa_edit.setMinimum(0.1)
        self.potência_ativa_edit.setMaximum(9999.9)
        self.potência_ativa_edit.setDecimals(1)
        self.comprimento_condutor_edit = QDoubleSpinBox()
        self.comprimento_condutor_edit.setMinimum(0.1)
        self.comprimento_condutor_edit.setMaximum(9999.9)
        self.comprimento_condutor_edit.setDecimals(1)
        self.seção_condutor_combo = QComboBox()
        self.seção_condutor_combo.addItems(['1.5', '2.5', '4', '6', '10', '16', '25', '35', '50', '70', '95', '120', '150'])
        self.material_cabo_combo = QComboBox()
        self.material_cabo_combo.addItems(['Cobre', 'Alumínio'])

        botao_calcular = QPushButton('Calcular')
        botao_calcular.clicked.connect(self.calcular_queda_tensão)

        grupo_entrada = QGroupBox('Dados de Entrada')
        layout_entrada = QGridLayout()
        layout_entrada.addWidget(QLabel('Nome:'), 0, 0)
        layout_entrada.addWidget(self.nome_edit, 0, 1)
        layout_entrada.addWidget(QLabel('Esquema de Ligação:'), 1, 0)
        layout_entrada.addWidget(self.esquema_ligação_combo, 1, 1)
        layout_entrada.addWidget(QLabel('Tensão de Alimentação (V):'), 2, 0)
        layout_entrada.addWidget(self.tensão_m_edit, 2, 1)
        layout_entrada.addWidget(QLabel('Fator de Potência:'), 3, 0)
        layout_entrada.addWidget(self.fator_potência_edit, 3, 1)
        layout_entrada.addWidget(QLabel('Potência Ativa (W):'), 4, 0)
        layout_entrada.addWidget(self.potência_ativa_edit, 4, 1)
        layout_entrada.addWidget(QLabel('Comprimento do Condutor (m):'), 5, 0)
        layout_entrada.addWidget(self.comprimento_condutor_edit, 5, 1)
        layout_entrada.addWidget(QLabel('Seção do Condutor (mm²):'), 6, 0)
        layout_entrada.addWidget(self.seção_condutor_combo, 6, 1)
        layout_entrada.addWidget(QLabel('Material do Cabo:'), 7, 0)
        layout_entrada.addWidget(self.material_cabo_combo, 7, 1)
        layout_entrada.addWidget(botao_calcular, 8, 0, 1, 2)

        grupo_entrada.setLayout(layout_entrada)

        return grupo_entrada


if __name__ == '__main__':
    app = QApplication(sys.argv)
    janela = CalculoQuedaTensão()
    janela.show()
    sys.exit(app.exec_())
