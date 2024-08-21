import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout
from PyQt5.QtGui import QPixmap, QIcon

class AdvancedInterface(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Interface Avançada')
        self.setGeometry(100, 100, 600, 400)

        layout = QVBoxLayout()

        # Criando um botão com uma imagem do Figma
        button = QPushButton()
        button.setIcon(QIcon('FEATURES/preview.jpg'))
        button.setIconSize(QPixmap('FEATURES/preview.jpg').size())
        
        layout.addWidget(button)

        self.setLayout(layout)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = AdvancedInterface()
    window.show()
    sys.exit(app.exec_())
