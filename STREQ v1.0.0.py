#STREQ - v1.0.0
import networkx as nx
import pandas as pd
import numpy as np
import xlsxwriter 
import warnings
import math
import xlrd
import sys
import os
from sklearn.metrics.pairwise import euclidean_distances
import datetime

#GUI Imports:
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfile 
from tkinter.filedialog import askdirectory
from tkinter import messagebox
from tkinter.ttk import Progressbar

#Global settings
warnings.filterwarnings("ignore")
np.set_printoptions(threshold=np.inf)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', -1)

##Global variables##
matrixName = ""
directory = "."
inter=False
barTick=0
overflow=False
excelMatrix=None
excelMatrix2=None
waitForMatrix = True
writer=None

##Helper functions##
def IsNotSquare(m): return not (m.shape[0] == m.shape[1])
def OpenFile(name): sys.stdout = open(name, "w")
def CloseFile(): sys.stdout.close()
def write_xlsx(table, sheet): table.replace(0, float("NaN")).to_excel(writer,sheet_name=sheet, header=True)    
def FormatSymmetrizeMatrix(matrix):
    for i in range(len(matrix)):
        for j in range(len(matrix)):
            if matrix.iloc[i,j] != 0:
                matrix.iloc[i,j]=np.round(matrix.iloc[i,j],3)
            if j>i:
                matrix.iloc[i,j] = 0 
def SymmetrizeMatrix(matrix):
    for i in range(len(matrix)):
        for j in range(len(matrix)):
            if j>i:
                matrix.iloc[i,j] = 0
def FormatMatrix(matrix):
    for i in range(len(matrix)):
        for j in range(len(matrix)):
            if matrix.iloc[i,j] !=0:
                matrix.iloc[i,j]=np.round(matrix.iloc[i,j],3)
def ClearMatrix(matrix):
    for i in range(len(matrix)):
        for j in range(len(matrix)):
            matrix.iloc[i,j]=0
def ShowError(title,message):
    global waitForMatrix
    global excelMatrix
    global excelMatrix2 
    global matrixName  

    matrixName="[None]"
    waitForMatrix=True
    excelMatrix=None  
    excelMatrix2=None  
    analyzeBtn.config(state=NORMAL)
    analyzeBtnText.set("Calculate")   
    matrixText.set("[None]")
    bar['value']=0   
    messagebox.showinfo(title,message)
    root.update() 
def OpenMatrix(filepath):
    global excelMatrix
    global excelMatrix2
    global inter
    global waitForMatrix

    excelMatrix=None
    excelMatrix2=None
    analyzeBtnText.set("Loading matrix file...")   
    isSquareBool = True
    inter = True    
    excelMatrix = pd.read_excel(filepath, sheet_name=0,index_col=0, header=0)
    excelMatrix=excelMatrix.fillna(0)
    labels=excelMatrix.index.tolist()
    if IsNotSquare(excelMatrix):                
        isSquareBool=False
    try:    
        excelMatrix2 = pd.read_excel(filepath, sheet_name=1,index_col=0, header=0)
        excelMatrix2=excelMatrix2.fillna(0) 
        if labels!=excelMatrix2.index.tolist():            
            ShowError("Input Error", "Nodes not matching between the 2 matrices!")  
            root.update() 
            return
        if IsNotSquare(excelMatrix2):
            isSquareBool=False    
    except (IndexError, ValueError):
        inter=False

    if isSquareBool:        
        analyzeBtn.config(state=NORMAL) 
    else:
        ShowError("Input Error", "Input matrix not square!") 
        root.update() 
        return

    waitForMatrix = False    
    if inter==False:
        analyzeBtnText.set("Calculate Intra")   
    else:
        analyzeBtnText.set("Calculate Inter")   
    root.update() 

#Inter Matrix calculations (2 matrices) 
def InterEuclidean():
    global excelMatrix
    global excelMatrix2
    size=len(excelMatrix)    

    IncrementBar("Calculating NetworkX part...")
    m1=np.transpose(excelMatrix)
    m2=np.transpose(excelMatrix2)
    try:      
        ed1=euclidean_distances(excelMatrix, excelMatrix2)
        ed2=euclidean_distances(m1, m2)   
    except ValueError:
        ShowError("Overflow","Numbers too big! Try scaling the input matrix down.")
        return

    IncrementBar("Calculating Euclidean Distances...")
    #Min Calculation
    a=[]
    aRow=[]
    aColumn=[]
    np.fill_diagonal(excelMatrix.values, excelMatrix.values.max())
    np.fill_diagonal(excelMatrix2.values, excelMatrix2.values.max())
    np.fill_diagonal(m1.values, m1.values.max())
    np.fill_diagonal(m2.values, m2.values.max())
    for i in range(size):
        min1=min(np.min(excelMatrix, axis=1)[i],np.min(excelMatrix2, axis=1)[i])
        min2=min(np.min(m1, axis=1)[i],np.min(m2, axis=1)[i])
        a.append(min1)
        a.append(min2)    
        aRow.append(min1)
        aColumn.append(min2)

    #Max Calculation
    b=[]
    bRow=[]
    bColumn=[]
    np.fill_diagonal(excelMatrix.values, excelMatrix.values.min())
    np.fill_diagonal(excelMatrix2.values, excelMatrix2.values.min())
    np.fill_diagonal(m1.values, m1.values.min())
    np.fill_diagonal(m2.values, m2.values.min())
    for i in range(size):
        max1=max(np.max(excelMatrix, axis=1)[i],np.max(excelMatrix2, axis=1)[i])
        max2=max(np.max(m1, axis=1)[i],np.max(m2, axis=1)[i])
        b.append(max1)
        b.append(max2)   
        bRow.append(max1)
        bColumn.append(max2)
     
    #Denominator Calculation
    c=[(size-1)*((y-x)**2) for x,y in zip(a,b)]
    cRow=[(size-1)*((y-x)**2) for x,y in zip(aRow,bRow)]
    cColumn=[(size-1)*((y-x)**2) for x,y in zip(aColumn,bColumn)]
    for x in c:   
        if x<0:
            ShowError("Overflow","Numbers too big! Try scaling the input matrix down.")
            return

    #Nominator Calculation
    d=[]
    rows=[]
    columns=[] 
    for i in range(size):
        tempA=ed1[i,i]**2
        tempB=ed2[i,i]**2
        if tempA<0 or tempB<0:
            ShowError("Overflow","Numbers too big! Try scaling the input matrix down.")
            return        
        d.append(tempA)
        d.append(tempB)
        rows.append(tempA)
        columns.append(tempB)

    #Printing Results
    lastsum=0
    lastsumW=0
    for i in range(len(d)):
        if c[i]!=0:
            lastsum= lastsum+d[i]
            lastsumW= lastsumW+d[i]/c[i]    
    if lastsum<0:
        ShowError("Overflow","Numbers too big! Try scaling the input matrix down.")
        return
    print("Euclidean Distance Absolute:", np.round(math.sqrt(lastsum),3))  
    print("Euclidean Distance Normalized:", np.round(math.sqrt(lastsumW)/math.sqrt(size*2),3))  
    print("")

    lastsum=0
    lastsumW=0
    for i in range(len(rows)):
        if cRow[i]!=0:
            lastsum= lastsum+rows[i]
            lastsumW= lastsumW+rows[i]/cRow[i]
    if lastsum<0:
        ShowError("Overflow","Numbers too big! Try scaling the input matrix down.")
        return
    print("Euclidean Distance Row Absolute:", np.round(math.sqrt(lastsum),3))  
    print("Euclidean Distance Row Normalized:", np.round(math.sqrt(lastsumW)/math.sqrt(size),3))    
    print("")

    lastsum=0
    lastsumW=0
    for i in range(len(columns)):
        if cColumn[i]!=0:
            lastsum= lastsum+columns[i]
            lastsumW= lastsumW+columns[i]/cColumn[i]
    if lastsum<0:
        ShowError("Overflow","Numbers too big! Try scaling the input matrix down.")
        return
    print("Euclidean Distance Column Absolute:", np.round(math.sqrt(lastsum),3))  
    print("Euclidean Distance Column Normalized:", np.round(math.sqrt(lastsumW)/math.sqrt(size),3))   
    CloseFile()
def InterJaccard():
    IncrementBar("Calculating Jaccard Matching...")
    root.update() 
    size=len(excelMatrix)
    num=0
    denW=0
    for x in range(size): 
        for y in range(size):  
            if x!=y:                        
                if excelMatrix.iloc[x,y]>excelMatrix2.iloc[x,y]:    #w
                    denW+=excelMatrix.iloc[x,y]                     #w
                else:                                               #w
                    denW+=excelMatrix2.iloc[x,y]                    #w    
                num+=min(excelMatrix.iloc[x,y],excelMatrix2.iloc[x,y])     
    print("Jaccard Matching Absolute:", num)     
    print("Jaccard Matching Normalized:", num/denW)   

    CloseFile()    
def InterSimple():
    IncrementBar("Calculating Simple Matching...")
    root.update() 
    size=len(excelMatrix)
    num=0
    denW=0
    for x in range(size): 
        for y in range(size): 
            if x!=y:               
                if excelMatrix.iloc[x,y]>excelMatrix2.iloc[x,y]:   
                    denW+=excelMatrix.iloc[x,y]                
                else:                                            
                    denW+=excelMatrix2.iloc[x,y]    
                if excelMatrix.iloc[x,y]==0 and excelMatrix2.iloc[x,y]==0:
                    denW+=1  
                    num+=1                                                     
                else:   
                    num+=min(excelMatrix.iloc[x,y],excelMatrix2.iloc[x,y])      
    print("Simple Matching Absolute:", num)     
    print("Simple Matching Normalized:", num/denW)    
    CloseFile()

#Intra Matrix calculations (1 matrices)
def IntraEuclidean():
    SMMatrix=excelMatrix.copy(True)
    SMAMatrix=excelMatrix.copy(True)
    ClearMatrix(SMMatrix)
    ClearMatrix(SMAMatrix)
    size=len(SMMatrix)

    #ROWS
    IncrementBar("Calculatig Row Matrix...")
    SMRMatrix=excelMatrix.copy(True)
    SMRAMatrix=excelMatrix.copy(True)
    ClearMatrix(SMRMatrix)
    ClearMatrix(SMRAMatrix)
    G = nx.from_pandas_adjacency(excelMatrix, create_using=nx.DiGraph) 

    edgesToAdd=[]
    for x in G.nodes():
        if G.out_degree(x)==0:
            edgesToAdd.append(x)
    for x in edgesToAdd:
        G.add_edge(x,x)        

    sEdges=list(G.edges())
    sNodes=list(G.nodes())
    sNodesLen=len(sNodes)
    sEdgesLen=len(sEdges)

    xNode=None
    xNodePred=None
    edgeList=[]
    tempList=[]
    for x in range(sEdgesLen):  #O(m)
        xNode=sEdges[x][0]
        if  xNodePred is not None and xNode != xNodePred:
            edgeList.append(tempList)
            tempList=[]
        tempList.append(sEdges[x])
        xNodePred=xNode
    edgeList.append(tempList)
    
    for x in range(sNodesLen): #O(n*m^2)
        for y in range(len(edgeList)):
            if y>x: 
                rowSum=0
                rowSumA=0
                maxEdge=0   
                minEdge=np.inf                                        
                nodeList=[]  
                smHelper=0                                 
                nX=sNodes[x]
                nY=sNodes[y]   
                for s,t in (edgeList[x]+edgeList[y]):
                    if s!=t:
                        try:
                            eXw=G[nX][t]["weight"]                                                  
                        except KeyError:
                            eXw=0    
                        try:
                            eYw=G[nY][t]["weight"]                                                 
                        except KeyError:
                            eYw=0     

                        if nY != t and nX!= t:   
                            if t not in nodeList:                   
                                nodeList.append(t)                 
                            else:                                   
                                smHelper+=1                                                   
                            if maxEdge<max(eXw,eYw):
                                maxEdge=max(eXw,eYw)               
                            if minEdge>min(eXw,eYw):
                                minEdge=min(eXw,eYw)   
                if len(nodeList)==smHelper:
                    minEdge=0
                nodeList=[] 
            
                for s,t in (edgeList[x]+edgeList[y]):
                    if s!=t:
                        try:
                            eXw=G[nX][t]["weight"]                                                  
                        except KeyError:
                            eXw=0    
                        try:
                            eYw=G[nY][t]["weight"]                                                 
                        except KeyError:
                            eYw=0           

                        if nY != t and nX!= t:                      
                            if t not in nodeList:                     
                                nodeList.append(t)  
                                tempA=(eYw-eXw)**2
                                tempB=(maxEdge-minEdge)**2                                
                                if tempA<0 or tempB<0:
                                    ShowError("Overflow","Number too big! Try scaling the input matrix down.")
                                    return
                                if tempB!=0 and (maxEdge-minEdge)!=0:
                                    rowSum+= tempA/((size-2)*tempB) 
                                    rowSumA+= tempA        

                SMRMatrix.iloc[y,x]=math.sqrt(rowSum)/math.sqrt(size-2)
                SMRAMatrix.iloc[y,x]=math.sqrt(rowSumA)
                SMMatrix.iloc[y,x]=rowSum
                SMAMatrix.iloc[y,x]=rowSumA

    #COLUMNS
    IncrementBar("Calculatig Column Matrix...")
    SMCMatrix=excelMatrix.copy(True)
    SMCAMatrix=excelMatrix.copy(True)
    ClearMatrix(SMCMatrix)
    ClearMatrix(SMCAMatrix)
    G=G.reverse()
    
    edgesToAdd=[]
    for x in G.nodes():
        if G.out_degree(x)==0:
            edgesToAdd.append(x)
    for x in edgesToAdd:
        G.add_edge(x,x)        

    sEdges=list(G.edges())
    sNodes=list(G.nodes())
    sNodesLen=len(sNodes)
    sEdgesLen=len(sEdges)

    xNode=None
    xNodePred=None
    edgeList=[]
    tempList=[]
    for x in range(sEdgesLen):  #O(m)
        xNode=sEdges[x][0]
        if  xNodePred is not None and xNode != xNodePred:
            edgeList.append(tempList)
            tempList=[]
        tempList.append(sEdges[x])
        xNodePred=xNode
    edgeList.append(tempList)

    for x in range(sNodesLen): #O(n*m^2)
        for y in range(len(edgeList)):
            if y>x: 
                rowSum=0
                rowSumA=0
                maxEdge=0   
                minEdge=np.inf                                        
                nodeList=[]  
                smHelper=0                                 
                nX=sNodes[x]
                nY=sNodes[y]   
                for s,t in (edgeList[x]+edgeList[y]):
                    if s!=t:
                        try:
                            eXw=G[nX][t]["weight"]                                                  
                        except KeyError:
                            eXw=0    
                        try:
                            eYw=G[nY][t]["weight"]                                                 
                        except KeyError:
                            eYw=0     

                        if nY != t and nX!= t:   
                            if t not in nodeList:                   
                                nodeList.append(t)                 
                            else:                                   
                                smHelper+=1                                                   
                            if maxEdge<max(eXw,eYw):
                                maxEdge=max(eXw,eYw)               
                            if minEdge>min(eXw,eYw):
                                minEdge=min(eXw,eYw)   
                if len(nodeList)==smHelper:
                    minEdge=0
                nodeList=[] 
            
                for s,t in (edgeList[x]+edgeList[y]):
                    if s!=t:
                        try:
                            eXw=G[nX][t]["weight"]                                                  
                        except KeyError:
                            eXw=0    
                        try:
                            eYw=G[nY][t]["weight"]                                                 
                        except KeyError:
                            eYw=0           

                        if nY != t and nX!= t:                      
                            if t not in nodeList:                     
                                nodeList.append(t)  
                                tempA=(eYw-eXw)**2
                                tempB=(maxEdge-minEdge)**2                                
                                if tempA<0 or tempB<0:
                                    ShowError("Overflow","Number too big! Try scaling the input matrix down.")
                                    return
                                if tempB!=0 and (maxEdge-minEdge)!=0:
                                    rowSum+= tempA/((size-2)*tempB)    
                                    rowSumA+= tempA         

                SMCMatrix.iloc[y,x]=math.sqrt(rowSum)/math.sqrt(size-2)  
                SMCAMatrix.iloc[y,x]=math.sqrt(rowSumA)
                SMMatrix.iloc[y,x]+=rowSum  
                SMMatrix.iloc[y,x]=math.sqrt(SMMatrix.iloc[y,x])/math.sqrt(2*(size-2))                 
                SMAMatrix.iloc[y,x]+=rowSumA                   
                SMAMatrix.iloc[y,x]=math.sqrt(SMAMatrix.iloc[y,x])

    IncrementBar("Saving matrices to disk...")
    write_xlsx(SMMatrix, "Normalized ED Matrix")
    SMMatrix=None
    write_xlsx(SMRMatrix, "Normalized ED Row Matrix")
    SMRMatrix=None
    write_xlsx(SMCMatrix, "Normalized ED Column Matrix")
    SMCMatrix=None

    write_xlsx(SMAMatrix, "Absolute ED Matrix")
    SMAMatrix=None
    write_xlsx(SMRAMatrix, "Absolute ED Row Matrix")
    SMRAMatrix=None
    write_xlsx(SMCAMatrix, "Absolute ED Column Matrix")
    SMCAMatrix=None
def IntraJaccard(): 
    SMMatrix=excelMatrix.copy(True)
    SMAMatrix=excelMatrix.copy(True)
    DenMatrix=excelMatrix.copy(True)
    DenNMatrix=excelMatrix.copy(True)
    ClearMatrix(SMMatrix)
    ClearMatrix(SMAMatrix)
    ClearMatrix(DenMatrix)
    ClearMatrix(DenNMatrix)
    size=len(SMMatrix)

    #ROWS
    IncrementBar("Calculatig Row Matrix...")
    SMRMatrix=excelMatrix.copy(True)
    SMRAMatrix=excelMatrix.copy(True)
    ClearMatrix(SMRMatrix)
    ClearMatrix(SMRAMatrix)

    G = nx.from_pandas_adjacency(excelMatrix, create_using=nx.DiGraph) 

    edgesToAdd=[]
    for x in G.nodes():
        if G.out_degree(x)==0:
            edgesToAdd.append(x)
    for x in edgesToAdd:
        G.add_edge(x,x)        

    sEdges=list(G.edges())
    sNodes=list(G.nodes())
    sNodesLen=len(sNodes)
    sEdgesLen=len(sEdges)

    xNode=None
    xNodePred=None
    edgeList=[]
    tempList=[]
    for x in range(sEdgesLen):  #O(m)
        xNode=sEdges[x][0]
        if  xNodePred is not None and xNode != xNodePred:
            edgeList.append(tempList)
            tempList=[]
        tempList.append(sEdges[x])
        xNodePred=xNode
    edgeList.append(tempList)
    
    for x in range(sNodesLen): #O(n*m^2)
        for y in range(len(edgeList)):
            if y>x: 
                rowSum=0
                maxEdge=0                                              
                nodeList=[]                                             
                nX=sNodes[x]
                nY=sNodes[y]         
                for s,t in (edgeList[x]+edgeList[y]):
                    if s!=t:
                        try:
                            eXw=G[nX][t]["weight"]                                                  
                        except KeyError:
                            eXw=0    
                        try:
                            eYw=G[nY][t]["weight"]                                                 
                        except KeyError:
                            eYw=0                               
                           
                        if nY != t and nX!= t:                      
                            if t not in nodeList:                     
                                nodeList.append(t)                       
                            if maxEdge<max(eXw,eYw):
                                maxEdge=max(eXw,eYw) 
                            minEdge=min(eXw,eYw)
                            rowSum+=minEdge  
                rowSum/=2 #it just works 
                remainder=len(nodeList)  
                temp=(maxEdge*remainder)
                if temp!=0:
                    SMRMatrix.iloc[y,x]=rowSum/(maxEdge*remainder)
                DenNMatrix.iloc[y,x]=remainder                                    
                SMRAMatrix.iloc[y,x]=rowSum
                SMMatrix.iloc[y,x]=rowSum
                SMAMatrix.iloc[y,x]=rowSum
                DenMatrix.iloc[y,x]=maxEdge  

    #COLUMNS
    IncrementBar("Calculatig Column Matrix...")
    SMCMatrix=excelMatrix.copy(True)
    SMCAMatrix=excelMatrix.copy(True)
    ClearMatrix(SMCMatrix)
    ClearMatrix(SMCAMatrix)
    G=G.reverse()
    
    edgesToAdd=[]
    for x in G.nodes():
        if G.out_degree(x)==0:
            edgesToAdd.append(x)
    for x in edgesToAdd:
        G.add_edge(x,x)        

    sEdges=list(G.edges())
    sNodes=list(G.nodes())
    sNodesLen=len(sNodes)
    sEdgesLen=len(sEdges)

    xNode=None
    xNodePred=None
    edgeList=[]
    tempList=[]
    for x in range(sEdgesLen):  #O(m)
        xNode=sEdges[x][0]
        if  xNodePred is not None and xNode != xNodePred:
            edgeList.append(tempList)
            tempList=[]
        tempList.append(sEdges[x])
        xNodePred=xNode
    edgeList.append(tempList)

    for x in range(sNodesLen): #O(n*m^2)
        for y in range(len(edgeList)):
            if y>x: 
                rowSum=0
                maxEdge=0                                              
                nodeList=[]                                                   
                nX=sNodes[x]
                nY=sNodes[y]         
                for s,t in (edgeList[x]+edgeList[y]):
                    if s!=t:
                        try:
                            eXw=G[nX][t]["weight"]                                                  
                        except KeyError:
                            eXw=0    
                        try:
                            eYw=G[nY][t]["weight"]                                                 
                        except KeyError:
                            eYw=0                               
                           
                        if nY != t and nX!= t:                      
                            if t not in nodeList:                         
                                nodeList.append(t)                        
                            if maxEdge<max(eXw,eYw):
                                maxEdge=max(eXw,eYw) 
                            minEdge=min(eXw,eYw)
                            rowSum+=minEdge    

                rowSum/=2 #it just works 
                remainder=len(nodeList)
                temp=(maxEdge*remainder)
                if temp!=0:
                    SMCMatrix.iloc[y,x]=rowSum/(maxEdge*remainder)
                DenNMatrix.iloc[y,x]+=remainder 
                SMCAMatrix.iloc[y,x]=rowSum
                SMMatrix.iloc[y,x]+=rowSum                
                SMAMatrix.iloc[y,x]+=rowSum
                if DenMatrix.iloc[y,x]<maxEdge:
                    DenMatrix.iloc[y,x]=maxEdge
                temp=(DenMatrix.iloc[y,x]*DenNMatrix.iloc[y,x])
                if temp!=0:
                    SMMatrix.iloc[y,x]=SMMatrix.iloc[y,x]/(DenMatrix.iloc[y,x]*DenNMatrix.iloc[y,x])  
                else:
                    SMMatrix.iloc[y,x]=np.nan
    
    IncrementBar("Saving matrices to disk...")
    write_xlsx(SMMatrix, "Normalized JM Matrix")
    SMMatrix=None
    write_xlsx(SMRMatrix, "Normalized JM Row Matrix")
    SMRMatrix=None
    write_xlsx(SMCMatrix, "Normalized JM Column Matrix")
    SMCMatrix=None

    write_xlsx(SMAMatrix, "Absolute JM Matrix")
    SMAMatrix=None
    write_xlsx(SMRAMatrix, "Absolute JM Row Matrix")
    SMRAMatrix=None
    write_xlsx(SMCAMatrix, "Absolute JM Column Matrix")
    SMCAMatrix=None
def IntraSimple():
    SMMatrix=excelMatrix.copy(True)
    SMAMatrix=excelMatrix.copy(True)
    DenMatrix=excelMatrix.copy(True)
    ClearMatrix(SMMatrix)
    ClearMatrix(SMAMatrix)
    ClearMatrix(DenMatrix)
    size=len(SMMatrix)

    #ROWS
    IncrementBar("Calculatig Row Matrix...")
    SMRMatrix=excelMatrix.copy(True)
    SMRAMatrix=excelMatrix.copy(True)
    ClearMatrix(SMRMatrix)
    ClearMatrix(SMRAMatrix)

    G = nx.from_pandas_adjacency(excelMatrix, create_using=nx.DiGraph) 

    edgesToAdd=[]
    for x in G.nodes():
        if G.out_degree(x)==0:
            edgesToAdd.append(x)
    for x in edgesToAdd:
        G.add_edge(x,x)        

    sEdges=list(G.edges())
    sNodes=list(G.nodes())
    sNodesLen=len(sNodes)
    sEdgesLen=len(sEdges)
    maxMatches=sNodesLen-2

    xNode=None
    xNodePred=None
    edgeList=[]
    tempList=[]
    for x in range(sEdgesLen):  #O(m)
        xNode=sEdges[x][0]
        if  xNodePred is not None and xNode != xNodePred:
            edgeList.append(tempList)
            tempList=[]
        tempList.append(sEdges[x])
        xNodePred=xNode
    edgeList.append(tempList)
    
    for x in range(sNodesLen): #O(n*m^2)
        for y in range(len(edgeList)):
            if y>x: 
                rowSum=0
                maxEdge=0   
                smHelper=0
                nodeList=[]        
                nX=sNodes[x]
                nY=sNodes[y]         
                for s,t in (edgeList[x]+edgeList[y]):
                    if s!=t:
                        try:
                            eXw=G[nX][t]["weight"]                                                  
                        except KeyError:
                            eXw=0    
                        try:
                            eYw=G[nY][t]["weight"]                                                 
                        except KeyError:
                            eYw=0                               
                           
                        if nY != t and nX!= t:                      
                            if t not in nodeList:                   
                                nodeList.append(t)                 
                            else:                                   
                                smHelper+=1                       
                            if maxEdge<max(eXw,eYw):
                                maxEdge=max(eXw,eYw) 
                            minEdge=min(eXw,eYw)
                            rowSum+=minEdge        
                if rowSum==0:                                       
                    rowSum=smHelper+(maxMatches-len(nodeList))      
                else:                                               
                    rowSum=rowSum/2+(maxMatches-len(nodeList))     

                temp = maxEdge*maxMatches
                if temp!=0:
                    SMRMatrix.iloc[y,x]=rowSum/temp
                SMRAMatrix.iloc[y,x]=rowSum
                SMMatrix.iloc[y,x]=rowSum
                SMAMatrix.iloc[y,x]=rowSum
                DenMatrix.iloc[y,x]=maxEdge  

    #COLUMNS
    IncrementBar("Calculatig Column Matrix...")
    SMCMatrix=excelMatrix.copy(True)
    SMCAMatrix=excelMatrix.copy(True)
    ClearMatrix(SMCMatrix)
    ClearMatrix(SMCAMatrix)
    G=G.reverse()
    
    edgesToAdd=[]
    for x in G.nodes():
        if G.out_degree(x)==0:
            edgesToAdd.append(x)
    for x in edgesToAdd:
        G.add_edge(x,x)        

    sEdges=list(G.edges())
    sNodes=list(G.nodes())
    sNodesLen=len(sNodes)
    sEdgesLen=len(sEdges)

    xNode=None
    xNodePred=None
    edgeList=[]
    tempList=[]
    for x in range(sEdgesLen):  #O(m)
        xNode=sEdges[x][0]
        if  xNodePred is not None and xNode != xNodePred:
            edgeList.append(tempList)
            tempList=[]
        tempList.append(sEdges[x])
        xNodePred=xNode
    edgeList.append(tempList)

    for x in range(sNodesLen): #O(n*m^2)
        for y in range(len(edgeList)):
            if y>x: 
                rowSum=0
                maxEdge=0   
                smHelper=0
                nodeList=[]        
                nX=sNodes[x]
                nY=sNodes[y]         
                for s,t in (edgeList[x]+edgeList[y]):
                    if s!=t:
                        try:
                            eXw=G[nX][t]["weight"]                                                  
                        except KeyError:
                            eXw=0    
                        try:
                            eYw=G[nY][t]["weight"]                                                 
                        except KeyError:
                            eYw=0                               
                           
                        if nY != t and nX!= t:                      
                            if t not in nodeList:                  
                                nodeList.append(t)                 
                            else:                                  
                                smHelper+=1                        
                            if maxEdge<max(eXw,eYw):
                                maxEdge=max(eXw,eYw) 
                            minEdge=min(eXw,eYw)
                            rowSum+=minEdge      
                if rowSum==0:                                       
                    rowSum=smHelper+(maxMatches-len(nodeList))      
                else:                                               
                    rowSum=rowSum/2+(maxMatches-len(nodeList))      

                temp = maxEdge*maxMatches
                if temp!=0:
                    SMCMatrix.iloc[y,x]=rowSum/temp
                SMCAMatrix.iloc[y,x]=rowSum
                SMMatrix.iloc[y,x]+=rowSum
                SMAMatrix.iloc[y,x]+=rowSum
                if DenMatrix.iloc[y,x]<maxEdge:
                    DenMatrix.iloc[y,x]=maxEdge
                temp=(DenMatrix.iloc[y,x]*2*maxMatches) 
                if temp!=0:
                    SMMatrix.iloc[y,x]=SMMatrix.iloc[y,x]/temp   
                else:
                    SMMatrix.iloc[y,x]=np.nan
    
    IncrementBar("Saving matrices to disk...")
    write_xlsx(SMMatrix, "Normalized SM Matrix")
    SMMatrix=None
    write_xlsx(SMRMatrix, "Normalized SM Row Matrix")
    SMRMatrix=None
    write_xlsx(SMCMatrix, "Normalized SM Column Matrix")
    SMCMatrix=None

    write_xlsx(SMAMatrix, "Absolute SM Matrix")
    SMAMatrix=None
    write_xlsx(SMRAMatrix, "Absolute SM Row Matrix")
    SMRAMatrix=None
    write_xlsx(SMCAMatrix, "Absolute SM Column Matrix")
    SMCAMatrix=None

#GUI code
def AnalyzeClick():
    global writer
    analyzeBtn.config(state=DISABLED)      
    GetProgreessBarTick()
   
    analyzeBtnText.set("Calculating...")

    if inter:
        if algorithm.get()=="Euclidean Distance":
            OpenFile(str(matrixName.rsplit( ".", 1 )[ 0 ]) +" ED Inter Results.txt")
            InterEuclidean()
        elif algorithm.get()=="Jaccard Matching":
            OpenFile(str(matrixName.rsplit( ".", 1 )[ 0 ]) +" JM Inter Results.txt")
            InterJaccard()
        elif algorithm.get()=="Simple Matching":
            OpenFile(str(matrixName.rsplit( ".", 1 )[ 0 ]) +" SM Inter Results.txt")
            InterSimple()
    else:
        if algorithm.get()=="Euclidean Distance":
            writer = pd.ExcelWriter(str(matrixName.rsplit( ".", 1 )[ 0 ]) +" ED Intra Results.xlsx", engine = 'xlsxwriter')  
            IntraEuclidean()
        elif algorithm.get()=="Jaccard Matching":
            writer = pd.ExcelWriter(str(matrixName.rsplit( ".", 1 )[ 0 ]) +" JM Intra Results.xlsx", engine = 'xlsxwriter')  
            IntraJaccard()
        elif algorithm.get()=="Simple Matching":
            writer = pd.ExcelWriter(str(matrixName.rsplit( ".", 1 )[ 0 ]) +" SM Intra Results.xlsx", engine = 'xlsxwriter')  
            IntraSimple()

        excelSaved=False
        while not excelSaved:
            try:
                writer.save()
                writer.close()
                excelSaved=True            
            except xlsxwriter.exceptions.FileCreateError:
                excelSaved=False        
                messagebox.showinfo("Output error", "Excel File already open! Close it and click OK")

    bar['value']=100   
    analyzeBtn.config(state=DISABLED)      
    IncrementBar("All done!")
def OpenFileClick():
    global directory
    global matrixName    
    bar['value']=0
    analyzeBtnText.set("Calculate")
    path = askopenfilename(initialdir = directory, title = "Choose a matrix file", filetypes =[("Excel Files", "*.xlsx *.xls")]  )  
    if path!="":
        matrixName = os.path.basename(path)
        directory = os.path.dirname(path)        
        OpenMatrix(path) 
        matrixText.set(matrixName)
    else:        
        matrixText.set("[None]")
        waitForMatrix=True
    CheckStatus()   
def GetProgreessBarTick():
    global barTick
    steps=1
    if inter:
        if algorithm.get()=="Euclidean Distance":
            steps+=2
        if algorithm.get()=="Jaccard Matching":
            steps+=1
        if algorithm.get()=="Simple Matching":
            steps+=1
    else:
        if algorithm.get()=="Euclidean Distance":
            steps+=3
        if algorithm.get()=="Jaccard Matching":
            steps+=3
        if algorithm.get()=="Simple Matching":
            steps+=3
    barTick=100/steps
def IncrementBar(string):
    bar['value']+= barTick
    analyzeBtnText.set(string)     
    root.update() 
def CheckStatus(*asdf):
    if waitForMatrix:
        analyzeBtn.config(state=DISABLED)   
    else:
        analyzeBtn.config(state=NORMAL)  
        if inter:       
            analyzeBtnText.set("Calculate Inter")    
        else:
            analyzeBtnText.set("Calculate Intra")  

    bar['value']=0   
def center(win):
    win.update_idletasks()
    width = win.winfo_width()
    frm_width = win.winfo_rootx() - win.winfo_x()
    win_width = width + 2 * frm_width
    height = win.winfo_height()
    titlebar_height = win.winfo_rooty() - win.winfo_y()
    win_height = height + titlebar_height + frm_width
    x = win.winfo_screenwidth() // 2 - win_width // 2
    y = win.winfo_screenheight() // 2 - win_height // 2
    win.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    win.deiconify()

root = Tk(  )
root.geometry("240x205") 
root.resizable(0, 0)
center(root)
root.title("STREQ v1.0.0")

matrixLabel = Label(root, text ="Input matrix:", width=30).grid(row=0, sticky=EW)
matrixText = StringVar()
matrixText.set("[None]")
outputBtn = Button(root, textvariable=matrixText, command = OpenFileClick, width=30).grid(row=1, padx = 10, sticky=EW)

matrixLabel = Label(root, text ="Algorithm:", width=30).grid(row=2, pady=(10,0), sticky=EW)
algorithm = StringVar()
algorithm.set("Euclidean Distance") 
algorithmMenu = OptionMenu(root, algorithm, "Euclidean Distance", "Jaccard Matching", "Simple Matching", command =CheckStatus)
algorithmMenu.grid(row=3, column=0, padx=10, sticky=EW)

matrixLabel = Label(root, text ="Progress:", width=30).grid(row=4, pady=(10,0), sticky=EW)
bar = Progressbar(root, length=100) 
bar.grid(row=5, column=0, padx=10, sticky=EW)

analyzeBtnText = StringVar()
analyzeBtnText.set("Calculate")

analyzeBtn = Button(root, textvariable=analyzeBtnText, width=30, command = AnalyzeClick)
analyzeBtn.grid(row=6,  column=0, padx=10, pady=(10,0), sticky=EW)
analyzeBtn.config(state=DISABLED)   

root.mainloop()