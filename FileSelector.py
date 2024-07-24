from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLabel

class FileSelector(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.layout = QVBoxLayout(self)
        
        self.selectButton = QPushButton('Select Excel Files', self)
        self.selectButton.clicked.connect(self.showDialog)
        
        self.filesLabel = QLabel(self)
        self.layout.addWidget(self.selectButton)
        self.layout.addWidget(self.filesLabel)
        
        self.setWindowTitle('Excel File Selector')

    def showDialog(self):
        fileNames, _ = QFileDialog.getOpenFileNames(self, 'Open files', '', "Excel files (*.xlsx *.xls)")
        self.filesLabel.setText("\n".join(fileNames))
        # You can call the function to process these files or handle them as needed here

if __name__ == '__main__':
    app = QApplication([])
    ex = FileSelector()
    ex.show()
    app.exec()
