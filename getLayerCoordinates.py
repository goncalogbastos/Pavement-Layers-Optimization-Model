import pandas as pd
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

def exportCoordinatesToExcel(coordinates: list, kms: list, side: str):
    df = pd.DataFrame(coordinates)
    df.insert(loc=0, column = 'km', value = kms)
    fileName = f'layerCoordinates_{side}.xlsx'
    isFile = os.path.exists(fileName)
    if isFile:
        try:
            os.remove(fileName)
        except Exception as e:
            print(e)
            exit()

    df.to_excel(fileName, index = False, header = False, sheet_name = side)
    
    isFile = os.path.exists(fileName)
    if isFile:
        print(f'\nThe file {fileName} containing the coordinates was created.')
        time.sleep(1)





def getInitialY_OffsetPosition(km: int):
    return (km - 43000)/25

def getYPosition(pointValues: list, k: int):
    return sum(pointValues[:k])


def getCoordinates(distances: list, P1: list, P2: list, P3: list, P4: list, km: int):
    coordinates = []

    isZeroP1 = all(value == 0 for value in P1)
    isZeroP2 = all(value == 0 for value in P2)
    isZeroP3 = all(value == 0 for value in P3)
    isZeroP4 = all(value == 0 for value in P4)

    ZerosP1 = P1.count(0)
    ZerosP2 = P2.count(0)
    ZerosP3 = P3.count(0)
    ZerosP4 = P4.count(0)

    Y = getInitialY_OffsetPosition(km)

    if isZeroP1:
        coordinates.append([0])

    elif (isZeroP1 == False) & (isZeroP4 == False) & (isZeroP2 == True) & (isZeroP3 == True):
        x1 = distances[0]
        x4 = distances[3]

        if ZerosP1 == 2:
            coordinates.append([
                (x1,Y),(x1,Y - getYPosition(P1,1)),(x4,Y),(x4,Y - getYPosition(P4,1)), # 1st layer 1-4
                (x1,Y - getYPosition(P1,1)),(x1,Y - getYPosition(P1,2)),(x4,Y - getYPosition(P4,1)),(x4,Y - getYPosition(P4,2)) # 2nd layer 1-4          
            ])

        elif ZerosP1 == 1:
            coordinates.append([
                (x1,Y),(x1,Y - getYPosition(P1,1)),(x4,Y),(x4,Y - getYPosition(P4,1)), # 1st layer 1-4
                (x1,Y - getYPosition(P1,1)),(x1,Y - getYPosition(P1,2)),(x4,Y - getYPosition(P4,1)),(x4,Y - getYPosition(P4,2)), # 2nd layer 1-4          
                (x1,Y - getYPosition(P1,2)),(x1,Y - getYPosition(P1,3)),(x4,Y - getYPosition(P4,2)),(x4,Y - getYPosition(P4,3)) # 3rd layer 1-4
            ])

        elif ZerosP1 == 0:
            coordinates.append([
                (x1,Y),(x1,Y - getYPosition(P1,1)),(x4,Y),(x4,Y - getYPosition(P4,1)), # 1st layer
                (x1,Y - getYPosition(P1,1)),(x1,Y - getYPosition(P1,2)),(x4,Y - getYPosition(P4,1)),(x4,Y - getYPosition(P4,2)), # 2nd layer 1-4          
                (x1,Y - getYPosition(P1,2)),(x1,Y - getYPosition(P1,3)),(x4,Y - getYPosition(P4,2)),(x4,Y - getYPosition(P4,3)), # 3rd layer 1-4
                (x1,Y - getYPosition(P1,3)),(x1,Y - getYPosition(P1,4)),(x4,Y - getYPosition(P4,3)),(x4,Y - getYPosition(P4,4)) # 4th layer 1-4
            ])
        else:
            pass

    elif (isZeroP1 == False) & (isZeroP4 == False) & (isZeroP2 == False) & (isZeroP3 == False):
        x1 = distances[0]
        x2 = distances[1]
        x3 = distances[2]
        x4 = distances[3]

        if ZerosP1 == 2:
            coordinates.append([
                (x1,Y),(x1,Y - getYPosition(P1,1)),(x2,Y),(x2,Y - getYPosition(P2,1)), # 1st layer 1-2
                (x3,Y),(x3,Y - getYPosition(P3,1)),(x4,Y),(x4,Y - getYPosition(P4,1)), # 1st layer 3-4

                (x1,Y - getYPosition(P1,1)),(x1,Y - getYPosition(P1,2)),(x2,Y - getYPosition(P2,1)),(x2,Y - getYPosition(P2,2)), # 2nd layer 1-2
                (x3,Y - getYPosition(P3,1)),(x3,Y - getYPosition(P3,2)),(x4,Y - getYPosition(P4,1)),(x4,Y - getYPosition(P4,2)) # 2nd layer 3-4                  
            ])

        elif ZerosP1 == 1:
            coordinates.append([
                (x1,Y),(x1,Y - getYPosition(P1,1)),(x2,Y),(x2,Y - getYPosition(P2,1)), # 1st layer 1-2
                (x3,Y),(x3,Y - getYPosition(P3,1)),(x4,Y),(x4,Y - getYPosition(P4,1)), # 1st layer 3-4

                (x1,Y - getYPosition(P1,1)),(x1,Y - getYPosition(P1,2)),(x2,Y - getYPosition(P2,1)),(x2,Y - getYPosition(P2,2)), # 2nd layer 1-2
                (x3,Y - getYPosition(P3,2)),(x3,Y - getYPosition(P4,2)),(x4,Y - getYPosition(P3,1)),(x4,Y - getYPosition(P4,2)), # 2nd layer 3-4        

                (x1,Y - getYPosition(P1,2)),(x1,Y - getYPosition(P1,3)),(x2,Y - getYPosition(P2,2)),(x2,Y - getYPosition(P2,3)), # 3rd layer 1-2
                (x3,Y - getYPosition(P3,2)),(x3,Y - getYPosition(P3,2)),(x4,Y - getYPosition(P4,1)),(x4,Y - getYPosition(P4,2)) # 3rd layer 3-4                     
            ])

        elif ZerosP1 == 0:
            coordinates.append([
                (x1,Y),(x1,Y - getYPosition(P1,1)),(x2,Y),(x2,Y - getYPosition(P2,1)), # 1st layer 1-2
                (x3,Y),(x3,Y - getYPosition(P3,1)),(x4,Y),(x4,Y - getYPosition(P4,1)), # 1st layer 3-4

                (x1,Y - getYPosition(P1,1)),(x1,Y - getYPosition(P1,2)),(x2,Y - getYPosition(P2,1)),(x2,Y - getYPosition(P2,2)), # 2nd layer 1-2
                (x3,Y - getYPosition(P3,2)),(x3,Y - getYPosition(P3,2)),(x4,Y - getYPosition(P4,1)),(x4,Y - getYPosition(P4,2)), # 2nd layer 3-4        

                (x1,Y - getYPosition(P1,2)),(x1,Y - getYPosition(P1,3)),(x2,Y - getYPosition(P2,2)),(x2,Y - getYPosition(P2,3)), # 3rd layer 1-2
                (x3,Y - getYPosition(P3,2)),(x3,Y - getYPosition(P3,3)),(x4,Y - getYPosition(P4,2)),(x4,Y - getYPosition(P4,3)), # 3rd layer 3-4

                (x1,Y - getYPosition(P1,3)),(x1,Y - getYPosition(P1,4)),(x2,Y - getYPosition(P2,3)),(x2,Y - getYPosition(P2,4)), # 4th layer 1-2
                (x3,Y - getYPosition(P3,3)),(x3,Y - getYPosition(P3,4)),(x4,Y - getYPosition(P4,3)),(x4,Y - getYPosition(P4,4)) # 4th layer 3-4   
            ])
        else:
            pass
    else:
        coordinates.append([0])

    return coordinates


def calculateCoordinates(distances: list, P1: list, P2: list, P3: list, P4: list, kms: list):
    coordinates = []
    print("")
    i=0
    for km in tqdm(kms, desc="Getting Coordinates...", ncols=100):
        coordinates.append(getCoordinates(distances[i], P1[i], P2[i], P3[i], P4[i], km))
        i += 1
        time.sleep(0.01)

    return coordinates



def main(filePath: str, side: str):
    df = readExcel(filePath)
    distances,P1,P2,P3,P4,kms = readExcelData(filePath, df)
    coordinates = calculateCoordinates(distances,P1,P2,P3,P4,kms)    
    exportCoordinatesToExcel(coordinates, kms, side)

    print("\nThe program will exit now.")
    time.sleep(3)  




if __name__ == "__main__":
    main('layerSolutions_left_Civil.xlsx', "left")



