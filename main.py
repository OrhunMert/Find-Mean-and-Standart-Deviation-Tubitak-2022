import pandas as pd
import numpy as np
import time
import xlsxwriter

def ReadFile(FileName):
    
    # FileName can be "1sn_1.xlsx","1sn_2.xlsx"...
    
    df = pd.read_excel(FileName)
    matrix =np.array(df)
    
    # matrix equals your datas in excel as two dimensional 
    
    return matrix
    
def findDimensions(matrix):
    
    # matrix --> your excel file's datas
    
    rowNum = len(matrix)
    colNum = len(matrix[0])
   
    return rowNum,colNum

def findMean(Array):
    
    # if you want to find to mean,you will give you an array.
    
    mean = np.mean(Array)
    
    return mean

def findStd(Array):
    
    # if you want to find to standart dev,you will give you an array.
    
    std = np.std(Array)
    
    return std

def writeExcel(mean_Matrix,std_Matrix,rowNum,colNum):
   
    
    workbook = xlsxwriter.Workbook('output.xlsx')
    worksheet_means = workbook.add_worksheet("means")
    worksheet_std = workbook.add_worksheet("std")
    
    
    for i in range(0,rowNum):
        
        for j,value in enumerate(mean_Matrix[i]):
                
            worksheet_means.write(i,j,value)
     
    for i in range(0,rowNum):
        
        for j,value in enumerate(std_Matrix[i]):
                
            worksheet_std.write(i,j,value)
            
    
    workbook.close()
    
    
def MeanAndStdMatrix(List_Matrix,rowNum,colNum):
    
    # List_Matrix --> You have a lot of excel file and they are appending a list.
    # this function find to mean and standart dev.
    
    mean_matrix = [[0 for j in range(colNum)] for i in range(rowNum)] 
    std_matrix = [[0 for j in range(colNum)] for i in range(rowNum)] 
  
    
    for i in range(0,rowNum):
        
        for j in range(0,colNum):
            
            print("\n")
            temp_meanAndStdMatrix = []
            
            for k in range(0,len(List_Matrix)):
                
                temp_Matrix = List_Matrix[k]
                temp_meanAndStdMatrix.append(temp_Matrix[i][j])
                
                mean_matrix[i][j] = findMean(temp_meanAndStdMatrix)
                std_matrix[i][j] = findStd(temp_meanAndStdMatrix)
                
                print(str(k+1)+".matrix "+str(i)+".row "+str(j)+".column")
                
                if k == len(List_Matrix) - 1:
                    print("temp_meanAndStdMatrix : "+str(temp_meanAndStdMatrix))
   
    
    return mean_matrix,std_matrix
    
def Main(List_TextFiles):
    
    # List_TextFiles --> ["1sn1_1.xlsx","1sn_2.xlsx",...]
    
    List_Matrix = []
    
    for i in range(0,len(List_TextFiles)):
        
        matrix = ReadFile(List_TextFiles[i])
        rowNum , colNum = findDimensions(matrix)
        List_Matrix.append(matrix)
        
        print(matrix)
        print("Row Number: "+str(rowNum)+"\nColumn Number: "+str(colNum))
        
    mean_matrix , std_matrix = MeanAndStdMatrix(List_Matrix,rowNum,colNum)
    writeExcel(mean_matrix,std_matrix,rowNum,colNum)
    
    print("\nProgram finished")

      
print("Program started\n")

#List_TextFiles = ["1sn_1.xlsx","1sn_2.xlsx","1sn_3.xlsx",
#                 "1sn_4.xlsx","1sn_5.xlsx","1sn_6.xlsx",
#                 "1sn_7.xlsx","1sn_8.xlsx","1sn_9.xlsx",
#                 "1sn_10.xlsx"]

List_TextFiles = ["1sn_1_az.xlsx" , "1sn_1_az.xlsx"]


start_time = time.time()

Main(List_TextFiles)

print(str(start_time - time.process_time())+" seconds")