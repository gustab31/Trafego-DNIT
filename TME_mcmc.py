# =====================================================
# imports

from pathlib import Path
import sqlite3 as lite
import sys
import pandas as pd
import numpy as np
import csv
from csv import reader
from tkinter import *
import math
import statistics as sta 
import scipy.stats as st
import os
import win32com.client
import statsmodels.formula.api as smf 
import seaborn as sns
# =====================================================
input_file_path = "C:\\Users\\gusta\\Dropbox\\PC\\Documents\\mestrado\\etapa_de_programacao\\programacao_normal_mcmc\\tme_mcmc\\time.csv"

# =====================================================
# Trabalhando com os dados da Dist Beta

data_1 = open("dist_beta_2.txt","r")

dataString_1 = data_1.read()
#print(dataString_1)

dataList_1 = dataString_1.split("\n")
#print(dataList_1)

# Transformar todos os elementos em floats
#   0    1       2
# ["1","3.32","5.43"] 
for i in range(0, len(dataList_1),1):
    dataList_1[i] = dataList_1[i].replace(",","")
    dataList_1[i] = float(dataList_1[i])

#print(dataList_1)
# =====================================================
# Trabalhando com os dados da Dist Gamma 

data_2 = open("dist_gamma_2.txt","r")

dataString_2 = data_2.read()
#print(dataString_2)

dataList_2 = dataString_2.split("\n")
#print(dataList_2)

# Transformar todos os elementos em floats
#   0    1       2
# ["1","3.32","5.43"] 
for i in range(0, len(dataList_2),1):
    dataList_2[i] = dataList_2[i].replace(",","")
    dataList_2[i] = float(dataList_2[i])

#print(dataList_2)
# =====================================================
# Trabalhando com os dados da Dist Normal  

data_3 = open("dist_normal_2.txt","r")

dataString_3 = data_3.read()
#print(dataString_3)

dataList_3 = dataString_3.split("\n")
#print(dataList_3)

# Transformar todos os elementos em floats
for i in range(0, len(dataList_3),1):
    dataList_3[i] = dataList_3[i].replace(",","")
    dataList_3[i] = float(dataList_3[i])

#print(dataList_3)
# =====================================================
# Agrupar os dados 
dist_beta = dataList_1
dist_gamma = dataList_2
dist_normal = dataList_3


dist_beta.sort()
dist_gamma.sort()
dist_normal.sort()

TME1 = dist_beta
TME2 = dist_gamma
TME3 = dist_normal

# =====================================================
#trabalhando com os dados
separador = ';'

z1 = []
z2 = []
z3 = []
z4 = []
with open(input_file_path, 'r', newline='') as csv_file:
    for line_number, content in enumerate(csv_file):
        if line_number:  # pula cabeçalho
            colunas = content.strip().split(separador)
            z1.append( float(colunas[0]) )
            z2.append( float(colunas[1]) )
            z3.append( float(colunas[2]) )
            z4.append( float(colunas[3]) )

z1_soma = (sum([i for i in z1 if isinstance(i, int) or isinstance(i, float)]))
z2_soma = (sum([i for i in z2 if isinstance(i, int) or isinstance(i, float)]))           
z3_soma = (sum([i for i in z3 if isinstance(i, int) or isinstance(i, float)]))
z4_soma = (sum([i for i in z4 if isinstance(i, int) or isinstance(i, float)]))
soma_z = z1_soma+z2_soma+z3_soma+z4_soma
media_z = (z1_soma+z2_soma+z3_soma+z4_soma)/4
# =====================================================
tempo_beta = (sum([i for i in TME1 if isinstance(i, int) or isinstance(i, float)]))
tempo_gamma = (sum([i for i in TME2 if isinstance(i, int) or isinstance(i, float)]))
tempo_normal = (sum([i for i in TME3 if isinstance(i, int) or isinstance(i, float)]))
# =====================================================
## Distribuicao Beta
fluxo_tempo1 = z1_soma*tempo_beta
fluxo_tempo2 = z2_soma*tempo_beta
fluxo_tempo3 = z3_soma*tempo_beta
fluxo_tempo4 = z4_soma*tempo_beta 

TIMER_1 = fluxo_tempo1/soma_z 
TIMER_2 = fluxo_tempo2/soma_z 
TIMER_3 = fluxo_tempo3/soma_z 
TIMER_4 = fluxo_tempo4/soma_z
TIMER_beta=(TIMER_1+TIMER_2+TIMER_3+TIMER_4)/4

# =====================================================
## Distribuicao Gamma
fluxo_tempo5 = z1_soma*tempo_gamma
fluxo_tempo6 = z2_soma*tempo_gamma
fluxo_tempo7 = z3_soma*tempo_gamma
fluxo_tempo8 = z4_soma*tempo_gamma 

TIMER_5 = fluxo_tempo5/soma_z 
TIMER_6 = fluxo_tempo6/soma_z 
TIMER_7 = fluxo_tempo7/soma_z 
TIMER_8 = fluxo_tempo8/soma_z
TIMER_gamma=(TIMER_5+TIMER_6+TIMER_7+TIMER_8)/4

# =====================================================
## Distribuicao Normal
fluxo_tempo9 = z1_soma*tempo_normal
fluxo_tempo10 = z2_soma*tempo_normal
fluxo_tempo11 = z3_soma*tempo_normal
fluxo_tempo12 = z4_soma*tempo_normal 

TIMER_9 = fluxo_tempo9/soma_z 
TIMER_10 = fluxo_tempo10/soma_z 
TIMER_11 = fluxo_tempo11/soma_z 
TIMER_12 = fluxo_tempo12/soma_z
TIMER_normal=(TIMER_9+TIMER_10+TIMER_11+TIMER_12)/4


# =====================================================
def nivel_servico(input1):  
    if input1 > 0 and input1 <= 10:
        return "A"
    elif input1 > 10 and input1 <= 20:
        return "B"
    elif input1 > 20 and input1 <= 30:
        return "C"
    elif input1 > 30 and input1 <= 45:
        return "D"
    elif input1 > 45:
        return "E"
    else:
        return "F" 
# =====================================================
print("Distribuicao beta")
print(TIMER_beta)
LOS_beta = nivel_servico(TIMER_beta)
print("Nível de serviço = %s" % LOS_beta) 
# =====================================================
print("Distribuicao gamma")
print(TIMER_gamma)
LOS_gamma = nivel_servico(TIMER_gamma)
print("Nível de serviço = %s" % LOS_gamma) 
# =====================================================
print("Distribuicao normal")
print(TIMER_normal)
LOS_normal = nivel_servico(TIMER_normal)
print("Nível de serviço = %s" % LOS_normal) 
