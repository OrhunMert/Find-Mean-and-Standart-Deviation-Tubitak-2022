import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

Filename = "output_LM.xlsx" # first row is header of excel file so if there are values the your first row, you must be empty first row.
sheetName = "std"
binValue = 30 # values of bin_edges intervals for Histogram. 

def readExcel(FileName , SheetName):
    
    df = pd.read_excel(FileName , sheet_name = SheetName , header = 0)
    array = np.array(df)
    
    return array

def findDimension(Matrix):
    
    # row Number , Column Number
    return len(Matrix) , len(Matrix[0])

def findMax(array):
    # your array can be multidimension array. This function find the max value of array.
    
    return np.max(array)

def calculateMatrixofMeans(array , Max_Value):
    
    # if you want to see histogram your mean values, you calculated new matrix with max values  
    
    rowNum , colNum = findDimension(array)
    
    matrix = [[0 for j in range(colNum)] for i in range(rowNum)]
    
    for i in range(0 , rowNum):
        
        for j in range(0 , colNum):
           
            matrix[i][j] = array[i][j]/Max_Value
            
    return matrix 

def Histogram(array , bins = 10):
    
    #hist --> freq  , bin_edges --> intervals
    
    hist, bin_edges = np.histogram(array , bins = bins)

    return hist , bin_edges

def HistogramPlot(array , bins = 10):
    
    plt.hist(array , bins = bins , ec = 'blue' )
    plt.title('Histogram with Binwidth = '+str(bins))
    plt.xlabel(str(sheetName)+' Values')
    plt.ylabel('Frequency')
    
    plt.show()

def Main(Filename , sheetName):
    
    Input_matrix = readExcel(Filename , sheetName)
    rowNum , colNum = findDimension(Input_matrix)
    
    if sheetName == "means":
        Max_Value = findMax(Input_matrix)
        matrix = calculateMatrixofMeans(Input_matrix, Max_Value)
        hist , bin_edges = Histogram(matrix , binValue)
        print("\nhist:\n"+str(hist))
        print("\nbin_edges:\n"+str(bin_edges))
        HistogramPlot(matrix , binValue)
    
    else:
        hist , bin_edges = Histogram(Input_matrix , binValue)
        print("\nhist:\n"+str(hist))
        print("\nbin_edges:\n"+str(bin_edges))
        HistogramPlot(Input_matrix , binValue)
        
    
Main(Filename , sheetName)