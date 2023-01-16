import pandas as pd
import time
import os
from decimal import Decimal as D
from tqdm import tqdm


def readExcel(file: str):
    print("\n############################################")
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
        print(
            f"\nThe file with the name {file} does not exist. Please, create a copy of the file with the data needed and put it in the current directory."
        )
        exit()


def readExcelData(file: str, df: pd.DataFrame):
    print("\n############################################")
    print(f"\nLoading data from file...")
    time.sleep(1)
    P1 = []
    P2 = []
    P3 = []
    P4 = []
    kms = []
    distances = []
    se = []
    M = []
    P = []
    
    try:
        for _, r in df.iterrows():
            row_distances = [r.L1, r.L2, r.L3, r.L4]
            row_P1 = [r.P1_ABGE1, r.P1_ABGE2, r.P1_ABGE3, r.P1_ABGE4]
            row_P2 = [r.P2_ABGE1, r.P2_ABGE2, r.P2_ABGE3, r.P2_ABGE4]
            row_P3 = [r.P3_ABGE1, r.P3_ABGE2, r.P3_ABGE3, r.P3_ABGE4]
            row_P4 = [r.P4_ABGE1, r.P4_ABGE2, r.P4_ABGE3, r.P4_ABGE4]
            row_kms = r.km
            row_se = r.SE
            row_M = r.M
            row_P = r.P

            distances.append(row_distances)
            P1.append(row_P1)
            P2.append(row_P2)
            P3.append(row_P3)
            P4.append(row_P4)
            kms.append(row_kms)
            se.append(row_se)
            M.append(row_M)
            P.append(row_P)

        print(f"Data successfully loaded.")        
        time.sleep(1)
        return distances, P1, P2, P3, P4, kms, se, M, P
    except:
        print(
            f"\nThe file {file} must have correct headers. Please, rename the headers accordingly."
        )
        exit()


def exportCoordinatesToExcel(coordinates: list, kms: list, side: str):
    df = pd.DataFrame(coordinates)
    df.insert(loc=0, column="km", value=kms)
    fileName = f"./output/layerCoordinates_{side}.xlsx"
    isFile = os.path.exists(fileName)
    if isFile:
        try:
            os.remove(fileName)
        except Exception as e:
            print(e)
            exit()

    df.to_excel(fileName, index=False, header=False, sheet_name=side)

    isFile = os.path.exists(fileName)
    if isFile:
        print(f"\nThe file {fileName} containing the coordinates was created.")
        time.sleep(1)


def getXPosition(km: int):
    pass

def getYPosition(km: int):
    pass


def getInitialY_OffsetPosition(km: int):
    return (km - 43000) / 25


def getYPosition(pointValues: list, k: int):
    return sum(pointValues[:k])


def getCoordinates(distances: list, P1: list, P2: list, P3: list, P4: list, km: int, se: float, X: float, Y: float):
    coordinates = []

    isZeroP1 = all(value == 0 for value in P1)
    isZeroP2 = all(value == 0 for value in P2)
    isZeroP3 = all(value == 0 for value in P3)
    isZeroP4 = all(value == 0 for value in P4)

    ZerosP1 = P1.count(0)
    ZerosP2 = P2.count(0)
    ZerosP3 = P3.count(0)
    ZerosP4 = P4.count(0)

    se = float(se)
    #Y = getInitialY_OffsetPosition(km)
    Y = Y - 0.28

    if isZeroP1:
        coordinates.append([0])

    elif (
        (isZeroP1 == False)
        & (isZeroP4 == False)
        & (isZeroP2 == True)
        & (isZeroP3 == True)
    ):
        x1 = distances[0]
        x4 = distances[3]

        if (x1 == 0) | (x4 == 0):
            coordinates.append([0])

        elif ZerosP1 == 2:
            coordinates.append(
                [
                    (X+x1, Y),
                    (X+x1, Y - getYPosition(P1, 1)),
                    (X+x4, Y + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 1) + se * abs(x4 - x1)),  # 1st layer 1-4
                    (X+x1, Y - getYPosition(P1, 1)),
                    (X+x1, Y - getYPosition(P1, 2)),
                    (X+x4, Y - getYPosition(P4, 1) + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 2) + + se * abs(x4 - x1)),  # 2nd layer 1-4
                ]
            )

        elif ZerosP1 == 1:
            coordinates.append(
                [
                    (X+x1, Y),
                    (X+x1, Y - getYPosition(P1, 1)),
                    (X+x4, Y + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 1) + se * abs(x4 - x1)),  # 1st layer 1-4
                    (X+x1, Y - getYPosition(P1, 1)),
                    (X+x1, Y - getYPosition(P1, 2)),
                    (X+x4, Y - getYPosition(P4, 1) + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 2) + se * abs(x4 - x1)),  # 2nd layer 1-4
                    (X+x1, Y - getYPosition(P1, 2)),
                    (X+x1, Y - getYPosition(P1, 3)),
                    (X+x4, Y - getYPosition(P4, 2) + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 3) + se * abs(x4 - x1)),  # 3rd layer 1-4
                ]
            )

        elif ZerosP1 == 0:
            coordinates.append(
                [
                    (X+x1, Y),
                    (X+x1, Y - getYPosition(P1, 1)),
                    (X+x4, Y + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 1) + se * abs(x4 - x1)),  # 1st layer
                    (X+x1, Y - getYPosition(P1, 1)),
                    (X+x1, Y - getYPosition(P1, 2)),
                    (X+x4, Y - getYPosition(P4, 1) + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 2) + se * abs(x4 - x1)),  # 2nd layer 1-4
                    (X+x1, Y - getYPosition(P1, 2)),
                    (X+x1, Y - getYPosition(P1, 3)),
                    (X+x4, Y - getYPosition(P4, 2) + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 3) + se * abs(x4 - x1)),  # 3rd layer 1-4
                    (X+x1, Y - getYPosition(P1, 3)),
                    (X+x1, Y - getYPosition(P1, 4)),
                    (X+x4, Y - getYPosition(P4, 3) + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 4) + se * abs(x4 - x1)),  # 4th layer 1-4
                ]
            )
        else:
            pass

    elif (
        (isZeroP1 == False)
        & (isZeroP4 == False)
        & (isZeroP2 == False)
        & (isZeroP3 == False)
    ):
        x1 = distances[0]
        x2 = distances[1]
        x3 = distances[2]
        x4 = distances[3]

        if ZerosP1 == 2:
            coordinates.append(
                [
                    (X+x1, Y),
                    (X+x1, Y - getYPosition(P1, 1)),
                    (X+x2, Y + se * abs(x2 - x1)),
                    (X+x2, Y - getYPosition(P2, 1) + se * abs(x2 - x1)),  # 1st layer 1-2
                    (X+x3, Y + se * abs(x3 - x1)),
                    (X+x3, Y - getYPosition(P3, 1) + se * abs(x3 - x1)),
                    (X+x4, Y + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 1) + se * abs(x4 - x1)),  # 1st layer 3-4
                    (X+x1, Y - getYPosition(P1, 1)),
                    (X+x1, Y - getYPosition(P1, 2)),
                    (X+x2, Y - getYPosition(P2, 1) + se * abs(x2 - x1)),
                    (X+x2, Y - getYPosition(P2, 2) + se * abs(x2 - x1)),  # 2nd layer 1-2
                    (X+x3, Y - getYPosition(P3, 1) + se * abs(x3 - x1)),
                    (X+x3, Y - getYPosition(P3, 2) + se * abs(x3 - x1)),
                    (X+x4, Y - getYPosition(P4, 1) + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 2) + se * abs(x4 - x1)),  # 2nd layer 3-4
                ]
            )

        elif ZerosP1 == 1:
            coordinates.append(
                [
                    (X+x1, Y),
                    (X+x1, Y - getYPosition(P1, 1)),
                    (X+x2, Y + se * abs(x2 - x1)),
                    (X+x2, Y - getYPosition(P2, 1) + se * abs(x2 - x1)),  # 1st layer 1-2
                    (X+x3, Y + se * abs(x3 - x1)),
                    (X+x3, Y - getYPosition(P3, 1) + se * abs(x3 - x1)),
                    (X+x4, Y + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 1) + se * abs(x4 - x1)),  # 1st layer 3-4
                    (X+x1, Y - getYPosition(P1, 1)),
                    (X+x1, Y - getYPosition(P1, 2)),
                    (X+x2, Y - getYPosition(P2, 1) + se * abs(x2 - x1)),
                    (X+x2, Y - getYPosition(P2, 2) + se * abs(x2 - x1)),  # 2nd layer 1-2
                    (X+x3, Y - getYPosition(P3, 1) + se * abs(x3 - x1)),
                    (X+x3, Y - getYPosition(P3, 2) + se * abs(x3 - x1)),
                    (X+x4, Y - getYPosition(P4, 1) + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 2) + se * abs(x4 - x1)),  # 2nd layer 3-4
                    (X+x1, Y - getYPosition(P1, 2)),
                    (X+x1, Y - getYPosition(P1, 3)),
                    (X+x2, Y - getYPosition(P2, 2) + se * abs(x2 - x1)),
                    (X+x2, Y - getYPosition(P2, 3) + se * abs(x2 - x1)),  # 3rd layer 1-2
                    (X+x3, Y - getYPosition(P3, 2) + se * abs(x3 - x1)),
                    (X+x3, Y - getYPosition(P3, 3) + se * abs(x3 - x1)),
                    (X+x4, Y - getYPosition(P4, 2) + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 3) + se * abs(x4 - x1)),  # 3rd layer 3-4
                ]
            )

        elif ZerosP1 == 0:
            coordinates.append(
                [
                    (X+x1, Y),
                    (X+x1, Y - getYPosition(P1, 1)),
                    (X+x2, Y + se * abs(x2 - x1)),
                    (X+x2, Y - getYPosition(P2, 1) + se * abs(x2 - x1)),  # 1st layer 1-2
                    (X+x3, Y + se * abs(x3 - x1)),
                    (X+x3, Y - getYPosition(P3, 1) + se * abs(x3 - x1)),
                    (X+x4, Y + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 1) + se * abs(x4 - x1)),  # 1st layer 3-4
                    (X+x1, Y - getYPosition(P1, 1)),
                    (X+x1, Y - getYPosition(P1, 2)),
                    (X+x2, Y - getYPosition(P2, 1) + se * abs(x2 - x1)),
                    (X+x2, Y - getYPosition(P2, 2) + se * abs(x2 - x1)),  # 2nd layer 1-2
                    (X+x3, Y - getYPosition(P3, 1) + se * abs(x3 - x1)),
                    (X+x3, Y - getYPosition(P3, 2) + se * abs(x3 - x1)),
                    (X+x4, Y - getYPosition(P4, 1) + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 2) + se * abs(x4 - x1)),  # 2nd layer 3-4
                    (X+x1, Y - getYPosition(P1, 2)),
                    (X+x1, Y - getYPosition(P1, 3)),
                    (X+x2, Y - getYPosition(P2, 2) + se * abs(x2 - x1)),
                    (X+x2, Y - getYPosition(P2, 3) + se * abs(x2 - x1)),  # 3rd layer 1-2
                    (X+x3, Y - getYPosition(P3, 2) + se * abs(x3 - x1)),
                    (X+x3, Y - getYPosition(P3, 3) + se * abs(x3 - x1)),
                    (X+x4, Y - getYPosition(P4, 2) + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 3) + se * abs(x4 - x1)),  # 3rd layer 3-4
                    (X+x1, Y - getYPosition(P1, 3)),
                    (X+x1, Y - getYPosition(P1, 4)),
                    (X+x2, Y - getYPosition(P2, 3) + se * abs(x2 - x1)),
                    (X+x2, Y - getYPosition(P2, 4) + se * abs(x2 - x1)),  # 4th layer 1-2
                    (X+x3, Y - getYPosition(P3, 3) + se * abs(x3 - x1)),
                    (X+x3, Y - getYPosition(P3, 4) + se * abs(x3 - x1)),
                    (X+x4, Y - getYPosition(P4, 3) + se * abs(x4 - x1)),
                    (X+x4, Y - getYPosition(P4, 4) + se * abs(x4 - x1)),  # 4th layer 3-4
                ]
            )
        else:
            pass
    else:
        coordinates.append([0])

    return coordinates


def calculateCoordinates(
    distances: list, P1: list, P2: list, P3: list, P4: list, kms: list, se: float, M: float, P: float
):
    coordinates = []
    print("")
    i = 0
    for km in tqdm(kms, desc="Getting Coordinates...", ncols=100):
        coordinates.append(getCoordinates(distances[i], P1[i], P2[i], P3[i], P4[i], km, se[i], M[i], P[i]))
        i += 1
        #time.sleep(0.001)

    return coordinates


def main(filePath: str, side: str):
    df = readExcel(filePath)
    distances, P1, P2, P3, P4, kms, se, M, P = readExcelData(filePath, df)    
    coordinates = calculateCoordinates(distances, P1, P2, P3, P4, kms, se, M, P)
    exportCoordinatesToExcel(coordinates, kms, side)


main("./src/layerSolutions_left_Civil.xlsx", "left")
main("./src/layerSolutions_right_Civil.xlsx", "right")
