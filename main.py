import pandas as pd
import numpy as np
import math
import os
import time
from decimal import Decimal as D


ESP_MIN_TOTAL_ABGE = 0.25
ESP_MAX_CAMADA_ABGE = 0.30
ESP_MIN_CAMADA_ABGE = 0.11

def checkRightSituation(d1,d2,d3,d4):
    if (d3==0) & (d2 == 0):
        return "1-4"
    else:
        return "1-2-3-4"

def checkLeftSituation(e4,e3,e2,e1):
    if ((e3==0) & (e2 == 0)) & ((e1 != 0) & (e4 != 0)):
        return "1-4"
    else:
        return "1-2-3-4"

def modulus(a, b):
    return D(str(a)) // D(str(b))

def reminder(a, b):
    return D(str(a)) % D(str(b))

def totalNumberOfLayers(a,b):
    return math.ceil(a/b)

def constantNumberOfLayers(a,b):
    return modulus(a,b)

def remainingTicknessLayer(a,b):
    return reminder(a,b)

def calculateNumberLayers(esp1,esp2,t):
    N1 = constantNumberOfLayers(esp1,t)
    R1 = remainingTicknessLayer(esp1,t)
    N4 = constantNumberOfLayers(esp2,t)
    R4 = remainingTicknessLayer(esp2,t)
    return N1,R1,N4,R4

def optimizeLayersLeft1_4(situtation,e4,e3,e2,e1):
    if situtation == "1-4":
        t = ESP_MIN_CAMADA_ABGE
        t_ = 0
        NT1 = 0
        NT4 = 1

        while NT1 != NT4:            
            NT1 = totalNumberOfLayers(e1,t)
            NT4 = totalNumberOfLayers(e4,t)
            if NT1 == NT4:
                N1,R1,N4,R4 = calculateNumberLayers(e1, e4, t)
                if (R1 < D(str(ESP_MIN_CAMADA_ABGE))) | (R4 < D(str(ESP_MIN_CAMADA_ABGE))):
                    NT1 = 0
                    NT4 = 1
                else:                            
                    return round(N1,2),round(R1,2),round(N4,2),round(R4,2),round(t,2),round(t_,2)           
            if t > ESP_MAX_CAMADA_ABGE:
                break
            t = t + 0.01

        if t > ESP_MAX_CAMADA_ABGE:
            print('*****************************')
            t = ESP_MIN_CAMADA_ABGE
            run = True
            while run:
                N1,R1,N4,R4 = calculateNumberLayers(e1, e4, t)
                N1_ = N1 - 1
                N4_ = N4 - 1
                N1 = 1
                N4 = 1
                R1 = float(N1_)*t + float(R1)
                R4 = float(N4_)*t + float(R4)

                if (ESP_MAX_CAMADA_ABGE < R1) | (R1 < ESP_MIN_CAMADA_ABGE) | (ESP_MAX_CAMADA_ABGE < R4) | (R4 < ESP_MIN_CAMADA_ABGE):
                    t = t + 0.01
                    if t > 0.3:
                        N1 = 1
                        R1 = e1/2
                        N4 = 1
                        R4 = e4/2
                        t = e1/2
                        t_ = e4/2                        
                        run = False                    
                else:
                    run = False  
            
        return round(N1,2),round(R1,2),round(N4,2),round(R4,2),round(t,2),round(t_,2)
    else:              
        return 0,0,0,0,0,0


def main():
    df = pd.read_excel("DATA.xlsx")
    df = df.fillna(0)
    df = df.apply(pd.to_numeric)

    points = []
    for i,r in df.iterrows():
        row_points = [r.e4, r.e3, r.e2, r.e1, r.d1, r.d2, r.d3, r.d4]
        points.append(row_points)

    solutions = []
    i=0
    for km in points: 
        leftSituation = checkLeftSituation(km[0],km[1],km[2],km[3])
        N1,R1,N4,R4,t,t_ = optimizeLayersLeft1_4(leftSituation,km[0],km[1],km[2],km[3])
        solutions.append([float(N1),float(R1),float(N4),float(R4),t,t_])
        print(i)
        i += 1
        
    
    sol_df = pd.DataFrame(solutions)
    try:
        os.remove('solutions.xlsx')
    except: 
        pass

    sol_df.to_excel('solutions.xlsx', index = False, header=False)
'''

def main():
    N1,R1,N4,R4,t = optimizeLayersLeft1_4("1-4",0.25,0,0,0.33)
    print(N1,R1,N4,R4,t)
'''



if __name__ == "__main__":
    main()