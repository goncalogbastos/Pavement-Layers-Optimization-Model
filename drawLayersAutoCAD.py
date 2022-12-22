import pandas as pd
import math
import time
import os
from decimal import Decimal as D
from tqdm import tqdm




def readExcel(file: str):
    print(f"\nChecking if the file {file} exists...")
    time.sleep(1)
    isFile = os.path.exists(file)
    if isFile:
        df = pd.read_excel(file)
        df = df.fillna(0)
        df = df.apply(pd.to_numeric)
        print(f"File successfully readed.")
        time.sleep(1)
        return df
    else:
        print(f'\nThe file with the name {file} does not exist. Please, create a copy of the file with the data needed and put it in the current directory.')
        exit()


def readExcelData(file: str, df: pd.DataFrame):
    print(f"\nLoading data from file...")
    time.sleep(1)
    P1 = []
    P2 = []
    P3 = []
    P4 = []
    kms = []
    distances = []
    try:
        for i,r in df.iterrows():
            row_distances = [r.L1, r.L2, r.L3, r.L4]
            row_P1 = [r.P1_ABGE1, r.P1_ABGE2, r.P1_ABGE3, r.P1_ABGE4]
            row_P2 = [r.P2_ABGE1, r.P2_ABGE2, r.P2_ABGE3, r.P2_ABGE4]
            row_P3 = [r.P3_ABGE1, r.P3_ABGE2, r.P3_ABGE3, r.P3_ABGE4]
            row_P4 = [r.P4_ABGE1, r.P4_ABGE2, r.P4_ABGE3, r.P4_ABGE4]
            row_kms = r.km

            distances.append(row_distances)
            P1.append(row_P1)
            P2.append(row_P2)
            P3.append(row_P3)
            P4.append(row_P4)            
            kms.append(row_kms)

        print(f"Data successfully loaded.")
        time.sleep(1)
        return distances, P1, P2, P3, P4, kms
    except:
        print(f"\nThe file {file} must have correct headers. Please, rename the headers accordingly.")
        exit()


def getYOffsetPosition(km: int):
    return (km - 43000)/25



def getCoordinates(distances: list, P1: list, P2: list, P3: list, P4: list, km: int):

    coordinates = []

    X_P1_ABGE1 = []
    Y_P1_ABGE1 = []
    X_P1_ABGE2 = []
    Y_P1_ABGE2 = []
    X_P1_ABGE3 = []
    Y_P1_ABGE3 = []
    X_P1_ABGE4 = []
    Y_P1_ABGE4 = []

    X_P2_ABGE1 = []
    Y_P2_ABGE1 = []
    X_P2_ABGE2 = []
    Y_P2_ABGE2 = []
    X_P2_ABGE3 = []
    Y_P2_ABGE3 = []
    X_P2_ABGE4 = []
    Y_P2_ABGE4 = []

    X_P3_ABGE1 = []
    Y_P3_ABGE1 = []
    X_P3_ABGE2 = []
    Y_P3_ABGE2 = []
    X_P3_ABGE3 = []
    Y_P3_ABGE3 = []
    X_P3_ABGE4 = []
    Y_P3_ABGE4 = []

    X_P4_ABGE1 = []
    Y_P4_ABGE1 = []
    X_P4_ABGE2 = []
    Y_P4_ABGE2 = []
    X_P4_ABGE3 = []
    Y_P4_ABGE3 = []
    X_P4_ABGE4 = []
    Y_P4_ABGE4 = []



    isZeroP1 = all(value == 0 for value in P1)
    isZeroP2 = all(value == 0 for value in P2)
    isZeroP3 = all(value == 0 for value in P3)
    isZeroP4 = all(value == 0 for value in P4)

    ZerosP1 = P1.count(0)
    ZerosP2 = P2.count(0)
    ZerosP3 = P3.count(0)
    ZerosP4 = P4.count(0)

    Y = getYOffsetPosition(km)

    if isZeroP1:
        coordinates.append([0]*32)

    elif (isZeroP1 == False) & (isZeroP4 == False) & (isZeroP2 == True) & (isZeroP3 == True):
        x1 = distances[0]
        x4 = distances[3]

        if ZerosP1 == 2:
            coordinates.append([(x1,Y),(x1,Y-P1[0]),(x4,Y),(x4,Y-P4[1]),(x1,Y-P1[0]),(x1,Y-P1[0]-P1[1]),(x4,Y-P4[0]),(x4,Y-P4[0]-P4[1])])

        elif ZerosP1 == 1:
            coordinates.append([(x1,Y),(x1,Y-P1[0]),(x4,Y),(x4,Y-P4[1]),(x1,Y-P1[0]),(x1,Y-P1[0]-P1[1]),(x4,Y-P4[0]),(x4,Y-P4[0]-P4[1])])








    
    return coordinates





filePath = 'layerSolutions_left_Civil.xlsx'

df = readExcel(filePath)
distances,P1,P2,P3,P4,kms = readExcelData(filePath, df)


