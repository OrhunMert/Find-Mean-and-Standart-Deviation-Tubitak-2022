import pandas as pd
import numpy as np
import time
import xlsxwriter
import sys

# inputFile_Num = 100 ; you must be writting your input File's number.E.g : input_1.txt ,input_2.txt . . . ,input_100.txt
inputFile_Num = 100

# Limunance or count's camera measurements. We are using to read the input files.E.g : 0,0001sn.input_1.txt or 0,1.input_1.txt
data_SecondsName = "0,01"

# You have two options. 'L' or 'C'

"""
L  --> limunance
C  --> count 
cL --> corrected limunance
"""
# We are using to create the output file. E.g : 0,0001_L_output.xlsx
measurements_type = 'cL'


def findMean(Array):
    
    # if you want to find to mean,you will give you an array.
    
    mean = np.mean(Array)
    
    return mean

def findStd(Array):
    
    # if you want to find to standart dev,you will give you an array.
    
    std = np.std(Array)
    
    return std


def ReadFileExcel(FileName):
    
    # FileName can be "1sn_1.xlsx","1sn_2.xlsx"...
    
    df = pd.read_excel(FileName)
    matrix =np.array(df)
    
    # matrix equals your datas in excel as two dimensional 
    
    return matrix
    
def findDimensionsExcel(matrix):
    
    # matrix --> your excel file's datas
    
    rowNum = len(matrix)
    colNum = len(matrix[0])
   
    return rowNum,colNum

def ReadFileText(file , Starting_rowIndex , rowNum , colNum):
    
    matrix = [[0 for j in range(colNum)] for i in range(rowNum)] 
    
    #StartingRow should be: first row = 1 , second row = 2.... for row number of filename.txt
    for j in range(1,Starting_rowIndex):
            file.readline() # we arrived row of values
    
    
    for i in range(0 , rowNum):
        
        line = file.readline()
        string_valuesArray = line.split()
        
        temp_StringValues = string_valuesArray
        temp_StringValues = [s.replace(",",".") for s in temp_StringValues] # float değerde "," yerine nokta olmalı.
        
        float_ValuesOfLine = [float(v) for v in temp_StringValues] # String to float 
        for j in range(0 , colNum):
            matrix[i][j] = float_ValuesOfLine[j]
    
    return matrix


def findDimensionText(file , Starting_rowIndex):
    
    lenLines = [line.strip("\n") for line in file if line != "\n"]
    rowNum = len(lenLines) - 2
    file.seek(0)
    
    count = 1
    
    for line in file:
        
        row_line = line.rstrip()
        
        if count == Starting_rowIndex:
            row_line = row_line.split()
            colNum = len(row_line)
            break
        
            
        count +=1
        
    file.seek(0)

    return rowNum , colNum

def writeExcel(mean_Matrix,std_Matrix,rowNum,colNum):
   
    
   workbook = xlsxwriter.Workbook(data_SecondsName+'_output_'+measurements_type+'.xlsx')
   worksheet_means = workbook.add_worksheet("means")
   worksheet_std = workbook.add_worksheet("std")
    
    
   for i in range(0,rowNum):
        
       for j,value in enumerate(mean_Matrix[i]):
                
           worksheet_means.write(i+1,j,value)
     
   for i in range(0,rowNum):
        
       for j,value in enumerate(std_Matrix[i]):
                
           worksheet_std.write(i+1,j,value)
            
    
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

# if input files are excel , You must use the MainExcel function.   
def MainExcel(List_TextFiles):
    
    # List_TextFiles --> ["1sn1_1.xlsx","1sn_2.xlsx",...]
    
    List_Matrix = []
    
    before_rowNum = 0
    before_colNum = 0
    
    
    for i in range(0,len(List_TextFiles)):
        
        matrix = ReadFileExcel(List_TextFiles[i])
        rowNum , colNum = findDimensionsExcel(matrix)
        List_Matrix.append(matrix)
        
        if i != 0:
            
            if before_rowNum != rowNum or before_colNum != colNum:
                
                print("\nERROR!!!\nInput File's dimension is wrong !!! Error input file index is "+str(i+1))
                sys.exit(0)
        
        before_rowNum = rowNum
        before_colNum = colNum
        
        print(matrix)
        print("Row Number: "+str(rowNum)+"\nColumn Number: "+str(colNum)+"\n")
        
    mean_matrix , std_matrix = MeanAndStdMatrix(List_Matrix,rowNum,colNum)
    writeExcel(mean_matrix,std_matrix,rowNum,colNum)
    
    print("\nProgram finished")


def MainText(inputFile_Num , data_SecondsName , measurements_type , Starting_rowIndex = 3) :
    
    FileName = ''
    
    List_Matrix = []
    
    before_rowNum = 0
    before_colNum = 0
    
    for i in range(1 , inputFile_Num+1):
        
        FileName = data_SecondsName+'_'+str(i)+'.txt'
        f = open(FileName , 'r')
        print(f.name)
        
        rowNum , colNum = findDimensionText(f , Starting_rowIndex)
        print("\nrowNum: "+str(rowNum)+" colNum: "+str(colNum))
        
        matrix = np.array(ReadFileText(f , Starting_rowIndex , rowNum , colNum))
        List_Matrix.append(matrix)
        
        
        if i != 1:
            
            if before_rowNum != rowNum or before_colNum != colNum:
                
                print("\nERROR!!!\nInput File's dimension is wrong !!! Error input file index is "+str(i))
                sys.exit(0)
                
        
        before_rowNum = rowNum
        before_colNum = colNum
        
        
    mean_matrix , std_matrix = MeanAndStdMatrix(List_Matrix , rowNum , colNum)
    writeExcel(mean_matrix,std_matrix,rowNum,colNum)
    
    f.close()
      
print("Program started\n")


start_time = time.time()


"""
--> You want to read the input text file, your text values format's important in files.
--> We are reading your file like this:
    input.txt
                       1            2        3
                  1  ushort
                  2   11          2035    2464 
    startingrow = 3 values......
                  .
                  .
                  .
"""

MainText(inputFile_Num , data_SecondsName , measurements_type , 3)  

print("--- %s seconds ---" % (time.time() - start_time))