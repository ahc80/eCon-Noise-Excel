import pandas as pd
import matplotlib.pyplot as plt
import os
import glob

def combineExcelFiles(folderPath):
    # Define the Excel file pattern
    filePattern = os.path.join(folderPath, '*.xlsx')
    # List all Excel files in the folder
    excelFiles = glob.glob(filePattern)

    # Combine all Excel files into one DataFrame
    combinedDf = pd.DataFrame()
    for file in excelFiles:
        df = pd.read_excel(file)
        combinedDf = pd.concat([combinedDf, df], ignore_index=True)

    # Save the combined DataFrame to a new Excel file
    combinedDf.to_excel('combinedExcel.xlsx', index=False)

    return combinedDf

def plotColumns(dataframe):
    # Plot each column as a separate graph
    for column in dataframe.columns:
        plt.figure()
        dataframe[column].plot(kind='line')
        plt.title(f'Graph for {column}')
        plt.ylabel(column)
        plt.xlabel('Index')
        plt.savefig(f'{column}.png')  # Save the plot as a PNG file
        plt.close()

# Use these functions
folderPath = 'path_to_your_folder_with_excels'
combinedDf = combineExcelFiles(folderPath)
plotColumns(combinedDf)
