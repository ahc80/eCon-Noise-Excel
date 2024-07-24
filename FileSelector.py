from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLabel, QHBoxLayout
import sys
import pandas as pd
import matplotlib.pyplot as plt
import os
import glob

class FileSelector(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        mainLayout = QVBoxLayout(self)

        self.selectButton = QPushButton('Select Folder with Excel Files', self)
        self.selectButton.clicked.connect(self.showDialog)

        self.filesLabel = QLabel('No folder selected', self)
        
        buttonLayout = QHBoxLayout()
        buttonLayout.addWidget(self.selectButton)
        
        labelLayout = QHBoxLayout()
        labelLayout.addWidget(self.filesLabel)

        mainLayout.addLayout(buttonLayout)
        mainLayout.addLayout(labelLayout)

        self.setLayout(mainLayout)
        self.setWindowTitle('Excel File Selector')
        self.setGeometry(500, 500, 400, 300)

    def showDialog(self):
        options = QFileDialog.Option
        folderPath = QFileDialog.getExistingDirectory(self, 'Select Folder', options=options)
        if folderPath:
            self.filesLabel.setText(folderPath)
            data = combineExcelFiles(folderPath)
            if not data.empty:
                plotColumns(data)
                self.filesLabel.setText('Processing Complete! Check the folder for results and combined Excel file.')
            else:
                self.filesLabel.setText('No Excel files found or combined data is empty.')

def combineExcelFiles(folderPath):
    filePattern = os.path.join(folderPath, '*.xlsx')
    excelFiles = glob.glob(filePattern)
    print(f"Files found: {excelFiles}")

    combinedDf = pd.DataFrame()

    if not excelFiles:
        print("No Excel files found. Please check the folder and file extensions.")
        return combinedDf

    for file in excelFiles:
        try:
            df = pd.read_excel(file)
            print(f"Data from {file}: {df.head()}")
            combinedDf = pd.concat([combinedDf, df], ignore_index=True)
        except Exception as e:
            print(f"Could not read {file}: {e}")

    combinedDf.to_excel('combinedExcel.xlsx', index=False)
    return combinedDf

def plotColumns(data):
    for column in data.columns:
        plt.figure()
        data[column].plot(kind='line')
        plt.title(f'Graph for {column}')
        plt.ylabel('Value')
        plt.xlabel('Index')
        plt.savefig(f'{column.replace("/", "_").replace(" ", "_")}.png')
        plt.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = FileSelector()
    ex.show()
    sys.exit(app.exec())
