import pandas as pd
import math
import time
import os
from decimal import Decimal as D
from tqdm import tqdm

ESP_MIN_TOTAL_ABGE = 0.25
ESP_MAX_CAMADA_ABGE = 0.30
ESP_MIN_CAMADA_ABGE = 0.12


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
    points = []
    kms = []
    try:
        for i, r in df.iterrows():
            row_points = [r.e4, r.e3, r.e2, r.e1, r.d1, r.d2, r.d3, r.d4]
            row_kms = r.km
            points.append(row_points)
            kms.append(row_kms)
        print(f"Data successfully loaded.")
        time.sleep(1)
        return points, kms
    except:
        print(
            f"\nThe file {file} must have at least the following headers: km, e4, e3, e2, e1, d1, d2, d3, d4. Please, rename the headers accordingly."
        )
        exit()


def exportSolutionsToExcel(solutions: list, kms: list, side: str):
    df = pd.DataFrame(solutions)
    df.insert(loc=0, column="km", value=kms)
    fileName = f"./output/layerSolutions_{side}.xlsx"
    isFile = os.path.exists(fileName)
    if isFile:
        try:
            os.remove(fileName)
        except Exception as e:
            print(e)
            exit()

    header = [
        "km",
        "P1_ABGE1",
        "P1_ABGE2",
        "P1_ABGE3",
        "P1_ABGE4",
        "P2_ABGE1",
        "P2_ABGE2",
        "P2_ABGE3",
        "P2_ABGE4",
        "P3_ABGE1",
        "P3_ABGE2",
        "P3_ABGE3",
        "P3_ABGE4",
        "P4_ABGE1",
        "P4_ABGE2",
        "P4_ABGE3",
        "P4_ABGE4",
    ]
    df.to_excel(fileName, index=False, header=header, sheet_name=side)

    isFile = os.path.exists(fileName)
    if isFile:
        print(f"\nThe file {fileName} containing the solutions was created.")
        time.sleep(1)


def checkSituation(a4, a3, a2, a1):
    if ((a3 == 0) & (a2 == 0)) & ((a1 != 0) & (a4 != 0)):
        return "1-4"
    elif ((a3 != 0) & (a2 != 0)) & ((a1 != 0) & (a4 != 0)):
        return "1-2-3-4"
    else:
        return "0-0-0-0"


def modulus(a, b):
    return D(str(a)) // D(str(b))


def reminder(a, b):
    return D(str(a)) % D(str(b))


def totalNumberOfLayers(a, b):
    return math.ceil(a / b)


def constantNumberOfLayers(a, b):
    return modulus(a, b)


def remainingTicknessLayer(a, b):
    return reminder(a, b)


def calculateNumberLayers(esp1, esp2, t):
    N = constantNumberOfLayers(esp1, t)
    R = remainingTicknessLayer(esp1, t)
    N_ = constantNumberOfLayers(esp2, t)
    R_ = remainingTicknessLayer(esp2, t)
    return N, R, N_, R_


def calculateNonConstantLayers(e4, e3, e2, e1):
    if (e2 == 0) & (e3 == 0):
        n = 2
        run = True
        while run:
            N1 = n - 1
            N4 = n - 1
            R1 = e1 / n
            R4 = e4 / n
            t = e1 / n
            t_ = e4 / n
            if (
                (round(R1, 2) >= ESP_MIN_CAMADA_ABGE)
                & (round(R4, 2) >= ESP_MIN_CAMADA_ABGE)
                & (round(t, 2) >= ESP_MIN_CAMADA_ABGE)
                & (round(t_, 2) >= ESP_MIN_CAMADA_ABGE)
            ):
               
                run = False
            else:
                n += 1
                
        return (
            round(N1, 2),
            round(R1, 2),
            round(N4, 2),
            round(R4, 2),
            round(t, 2),
            round(t_, 2),
        )
    else:
        pass


def convertTwoPointsToFinalLayers(N1, R1, N4, R4, t, t_):

    P1_ABGE1 = P1_ABGE2 = P1_ABGE3 = P1_ABGE4 = 0
    P2_ABGE1 = P2_ABGE2 = P2_ABGE3 = P2_ABGE4 = 0
    P3_ABGE1 = P3_ABGE2 = P3_ABGE3 = P3_ABGE4 = 0
    P4_ABGE1 = P4_ABGE2 = P4_ABGE3 = P4_ABGE4 = 0

    if t_ == 0:

        if (N1 == 1) & (N4 == 1):
            P1_ABGE1 = P4_ABGE1 = t
            P1_ABGE2 = R1
            P4_ABGE2 = R4
        elif (N1 == 2) & (N4 == 2):
            P1_ABGE1 = P4_ABGE1 = P1_ABGE2 = P4_ABGE2 = t
            P1_ABGE3 = R1
            P4_ABGE3 = R4
        elif (N1 == 3) & (N4 == 3):
            P1_ABGE1 = P4_ABGE1 = P1_ABGE2 = P4_ABGE2 = P1_ABGE3 = P4_ABGE3 = t
            P1_ABGE4 = R1
            P4_ABGE4 = R4
        else:
            pass
    else:

        if (N1 == 1) & (N4 == 1):
            P1_ABGE1 = t
            P4_ABGE1 = t_
            P1_ABGE2 = R1
            P4_ABGE2 = R4
        elif (N1 == 2) & (N4 == 2):
            P1_ABGE1 = P1_ABGE2 = t
            P4_ABGE1 = P4_ABGE2 = t_
            P1_ABGE3 = R1
            P4_ABGE3 = R4
        elif (N1 == 3) & (N4 == 3):
            P1_ABGE1 = P1_ABGE2 = P1_ABGE3 = t
            P4_ABGE1 = P4_ABGE2 = P4_ABGE3 = t_
            P1_ABGE4 = R1
            P4_ABGE4 = R4
        else:
            pass
    layers = [
        P1_ABGE1,
        P1_ABGE2,
        P1_ABGE3,
        P1_ABGE4,
        P2_ABGE1,
        P2_ABGE2,
        P2_ABGE3,
        P2_ABGE4,
        P3_ABGE1,
        P3_ABGE2,
        P3_ABGE3,
        P3_ABGE4,
        P4_ABGE1,
        P4_ABGE2,
        P4_ABGE3,
        P4_ABGE4,
    ]
    return layers


def convertFourPointsToFinalLayers(N1, R1, N2, R2, tA, tB, N3, R3, N4, R4, tC, tD):

    P1_ABGE1 = P1_ABGE2 = P1_ABGE3 = P1_ABGE4 = 0
    P2_ABGE1 = P2_ABGE2 = P2_ABGE3 = P2_ABGE4 = 0
    P3_ABGE1 = P3_ABGE2 = P3_ABGE3 = P3_ABGE4 = 0
    P4_ABGE1 = P4_ABGE2 = P4_ABGE3 = P4_ABGE4 = 0

    if tB == 0:
        if (N1 == 1) & (N2 == 1):
            P1_ABGE1 = P2_ABGE1 = tA
            P1_ABGE2 = R1
            P2_ABGE2 = R2
        elif (N1 == 2) & (N2 == 2):
            P1_ABGE1 = P2_ABGE1 = P1_ABGE2 = P2_ABGE2 = tA
            P1_ABGE3 = R1
            P2_ABGE3 = R2
        elif (N1 == 3) & (N2 == 3):
            P1_ABGE1 = P2_ABGE1 = P1_ABGE2 = P2_ABGE2 = P1_ABGE3 = P2_ABGE3 = tA
            P1_ABGE4 = R1
            P2_ABGE4 = R2
        else:
            pass
    else:
        if (N1 == 1) & (N2 == 1):

            P1_ABGE1 = tA
            P2_ABGE1 = tB
            P1_ABGE2 = R1
            P2_ABGE2 = R2

        elif (N1 == 2) & (N2 == 2):
            P1_ABGE1 = P1_ABGE2 = tA
            P2_ABGE1 = P2_ABGE2 = tB
            P1_ABGE3 = R1
            P2_ABGE3 = R2
        elif (N1 == 3) & (N2 == 3):
            P1_ABGE1 = P1_ABGE2 = P1_ABGE3 = tA
            P2_ABGE1 = P2_ABGE2 = P2_ABGE3 = tB
            P1_ABGE4 = R1
            P2_ABGE4 = R2
        else:
            pass

    if tD == 0:
        if (N3 == 1) & (N4 == 1):
            P3_ABGE1 = P4_ABGE1 = tC
            P3_ABGE2 = R3
            P4_ABGE2 = R4
        elif (N3 == 2) & (N4 == 2):
            P3_ABGE1 = P4_ABGE1 = P3_ABGE2 = P4_ABGE2 = tC
            P3_ABGE3 = R3
            P4_ABGE3 = R4
        elif (N3 == 3) & (N4 == 3):
            P3_ABGE1 = P4_ABGE1 = P3_ABGE2 = P4_ABGE2 = P3_ABGE3 = P4_ABGE3 = tC
            P3_ABGE4 = R3
            P4_ABGE4 = R4
        else:
            pass
    else:
        if (N3 == 1) & (N4 == 1):
            P3_ABGE1 = tC
            P4_ABGE1 = tD
            P3_ABGE2 = R3
            P4_ABGE2 = R4
        elif (N3 == 2) & (N4 == 2):
            P3_ABGE1 = P4_ABGE2 = tC
            P3_ABGE1 = P4_ABGE2 = tD
            P3_ABGE3 = R3
            P4_ABGE3 = R4
        elif (N3 == 3) & (N4 == 3):
            P3_ABGE1 = P3_ABGE2 = P3_ABGE3 = tC
            P4_ABGE1 = P4_ABGE2 = P4_ABGE3 = tD
            P3_ABGE4 = R3
            P4_ABGE4 = R4
        else:
            pass

    layers = [
        P1_ABGE1,
        P1_ABGE2,
        P1_ABGE3,
        P1_ABGE4,
        P2_ABGE1,
        P2_ABGE2,
        P2_ABGE3,
        P2_ABGE4,
        P3_ABGE1,
        P3_ABGE2,
        P3_ABGE3,
        P3_ABGE4,
        P4_ABGE1,
        P4_ABGE2,
        P4_ABGE3,
        P4_ABGE4,
    ]
    return layers


def optimizeLayersTwoPoints(e4, e3, e2, e1):

    t = ESP_MIN_CAMADA_ABGE
    t_ = 0
    NT1 = 0
    NT4 = 1

    while NT1 != NT4:
        NT1 = totalNumberOfLayers(e1, t)
        NT4 = totalNumberOfLayers(e4, t)
        if NT1 == NT4:
            N1, R1, N4, R4 = calculateNumberLayers(e1, e4, t)
            if (round(R1, 2) < D(str(ESP_MIN_CAMADA_ABGE))) | (
                round(R4, 2) < D(str(ESP_MIN_CAMADA_ABGE))
            ):
                NT1 = 0
                NT4 = 1
            else:
                return (
                    round(N1, 2),
                    round(R1, 2),
                    round(N4, 2),
                    round(R4, 2),
                    round(t, 2),
                    round(t_, 2),
                )
        if t > ESP_MAX_CAMADA_ABGE:
            break
        t = t + 0.01

    if t > ESP_MAX_CAMADA_ABGE:
        t = ESP_MIN_CAMADA_ABGE
        run = True
        while run:
            N1, R1, N4, R4 = calculateNumberLayers(e1, e4, t)
            N1_ = N1 - 1
            N4_ = N4 - 1
            N1 = N4 = 1
            R1 = float(N1_) * t + float(R1)
            R4 = float(N4_) * t + float(R4)

            if (
                (ESP_MAX_CAMADA_ABGE < round(R1, 2))
                | (round(R1, 2) < ESP_MIN_CAMADA_ABGE)
                | (ESP_MAX_CAMADA_ABGE < round(R4, 2))
                | (round(R4, 2) < ESP_MIN_CAMADA_ABGE)
            ):
                t = t + 0.01
                if t > 0.3:
                    N1, R1, N4, R4, t, t_ = calculateNonConstantLayers(e4, e3, e2, e1)
                    run = False
            else:
                run = False

    return (
        round(N1, 2),
        round(R1, 2),
        round(N4, 2),
        round(R4, 2),
        round(t, 2),
        round(t_, 2),
    )


def optimizeLayersFourPoints(e4, e3, e2, e1):
    N1, R1, N2, R2, tA, tB = optimizeLayersTwoPoints(e2, 0, 0, e1)
    N3, R3, N4, R4, tC, tD = optimizeLayersTwoPoints(e4, 0, 0, e3)
    return (
        round(N1, 2),
        round(R1, 2),
        round(N2, 2),
        round(R2, 2),
        round(tA, 2),
        round(tB, 2),
        round(N3, 2),
        round(R3, 2),
        round(N4, 2),
        round(R4, 2),
        round(tC, 2),
        round(tD, 2),
    )


def analyseLayers(points: list, kms: list):
    leftSolutions = []
    rightSolutions = []
    i = 0

    print("")
    for p in tqdm(points, desc="Analysing layers...", ncols=100):        
        leftSituation = checkSituation(p[0], p[1], p[2], p[3])
        rightSituation = checkSituation(p[7], p[6], p[5], p[4])

        if leftSituation == "1-4":
            N1, R1, N4, R4, t, t_ = optimizeLayersTwoPoints(p[0], p[1], p[2], p[3])
            layers = convertTwoPointsToFinalLayers(N1, R1, N4, R4, t, t_)
            layers_ = []
            for layer in layers:
                layers_.append(float(layer))
            leftSolutions.append(layers_)

        elif leftSituation == "1-2-3-4":
            (
                N1,
                R1,
                N2,
                R2,
                tA,
                tB,
                N3,
                R3,
                N4,
                R4,
                tC,
                tD,
            ) = optimizeLayersFourPoints(p[0], p[1], p[2], p[3])
            layers = convertFourPointsToFinalLayers(
                N1, R1, N2, R2, tA, tB, N3, R3, N4, R4, tC, tD
            )
            layers_ = []
            for layer in layers:
                layers_.append(float(layer))
            leftSolutions.append(layers_)
        else:
            leftSolutions.append([0] * 16)

        if rightSituation == "1-4":
            N1, R1, N4, R4, t, t_ = optimizeLayersTwoPoints(p[7], p[6], p[5], p[4])
            layers = convertTwoPointsToFinalLayers(N1, R1, N4, R4, t, t_)
            layers_ = []
            for layer in layers:
                layers_.append(float(layer))
            rightSolutions.append(layers_)

        elif rightSituation == "1-2-3-4":
            (
                N1,
                R1,
                N2,
                R2,
                tA,
                tB,
                N3,
                R3,
                N4,
                R4,
                tC,
                tD,
            ) = optimizeLayersFourPoints(p[7], p[6], p[5], p[4])
            layers = convertFourPointsToFinalLayers(
                N1, R1, N2, R2, tA, tB, N3, R3, N4, R4, tC, tD
            )
            layers_ = []
            for layer in layers:
                layers_.append(float(layer))
            rightSolutions.append(layers_)
        else:
            rightSolutions.append([0] * 16)

        i += 1
        #time.sleep(0.001)

    return leftSolutions, rightSolutions


def main(file: str):
    df = readExcel(file)
    points, kms = readExcelData(file, df)
    leftSolutions, rightSolutions = analyseLayers(points, kms)
    exportSolutionsToExcel(leftSolutions, kms, "left")
    exportSolutionsToExcel(rightSolutions, kms, "right")


if __name__ == "__main__":
    main("./src/DATA.xlsx")
    import copy_excel_columns
    import coordinates

    print("\nThe program will exit now.")
    time.sleep(3)
