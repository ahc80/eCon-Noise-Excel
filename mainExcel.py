import pandas as pd
import matplotlib.pyplot as plt
import os
import glob

def combineExcelFiles(folderPath):
    filePattern = os.path.join(folderPath, '*.xlsx')  # Ensure this matches your file extension
    print(f"Looking for files with pattern: {filePattern}")  # Debug print to check the file pattern
    
    excelFiles = glob.glob(filePattern)
    print(f"Files found: {excelFiles}")  # Debug print to check the files list

    combinedDf = pd.DataFrame()

    if not excelFiles:
        print("No Excel files found. Please check the folder and file extensions.")
        return combinedDf

    for file in excelFiles:
        df = pd.read_excel(file)
        print(f"Data from {file}: {df.head()}")  # Debug print to show data being read
        combinedDf = pd.concat([combinedDf, df], ignore_index=True)

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

# Integrate these functions into your existing PyQt6 setup for GUI interaction
