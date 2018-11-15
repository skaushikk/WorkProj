# -*- coding: utf-8 -*-
"""
Created on Fri Sep  2 14:36:10 2016

@author: ksirvole
"""
import csv, time
import math
import os
import sys
from subprocess import Popen
import OrcFxAPI
import numpy as np
import xlwings as xw
from shutil import copyfile
from scipy.optimize import fsolve

def PEMS_Loads(filename, winchx):
    model = OrcFxAPI.Model(filename)
    pipelineWinch = model['Winch1']
    pipeline = model['Pipeline']
    
#    pipelineWinchX = pipelineWinch.ConnectionX[0]
    pipelineWinch.ConnectionX[0] = winchx
    
    model.CalculateStatics()                         
    pipeline_tension = pipeline.StaticResult('Effective Tension', OrcFxAPI.oeEndA)
    pipeline_bending_moment = pipeline.StaticResult('y-Bend Moment', OrcFxAPI.oeEndA)
    pipelineWinchTension = pipelineWinch.StageValue[0]
#    model.SaveData(filename)
    return pipeline_tension, pipeline_bending_moment,pipelineWinchTension

def PEMS_BM(filename, winchx, target_tension = 0):
    model = OrcFxAPI.Model(filename)
    pipelineWinch = model['Winch1']
    pipeline = model['Pipeline']
    
#    pipelineWinchX = pipelineWinch.ConnectionX[0]
    pipelineWinch.ConnectionX[0] = winchx
    
    model.CalculateStatics()                         
    pipeline_tension = pipeline.StaticResult('Effective Tension', OrcFxAPI.oeEndA)
    pipeline_bending_moment = pipeline.StaticResult('y-Bend Moment', OrcFxAPI.oeEndA)
    
    TensionDiff = target_tension - pipeline_tension
    pipelineWinchTension = pipelineWinch.StageValue[0]
    pipelineWinch.StageValue[0] = pipelineWinchTension + TensionDiff
    
    model.CalculateStatics()
    pipeline_tension = pipeline.StaticResult('Effective Tension', OrcFxAPI.oeEndA)
    pipeline_bending_moment = pipeline.StaticResult('y-Bend Moment', OrcFxAPI.oeEndA)                        
#    pipelineWinchTension = pipelineWinch.StageValue[0]
    model.SaveData(filename)
    return pipeline_bending_moment



def PEMS(Fileinfo, Tloads, PipeEndA, WinchB, Contents, Tol):
    # Set the input file based on contents
    global PEMS_Base_File
    if Contents == 'Flooded':
        PEMS_Base_File = 'PEMS_base_F.dat'
    if Contents == 'Empty':
        PEMS_Base_File = 'PEMS_base_E.dat'

    basefilename = Fileinfo[0] + "\\" + PEMS_Base_File
    tarfilename = Fileinfo[1]                       
    copyfile (basefilename, tarfilename)                       
    
    model = OrcFxAPI.Model(tarfilename)
    
    
    ## Define pipeline and winch end coordinates

    # Grab an individual Winch and pipeline
    pipelineWinch = model['Winch1']
    pipeline = model['Pipeline']

    # Assign pipe endA coordinates
    pipeline.EndAX = PipeEndA[0]
    pipeline.EndAZ = PipeEndA[2]
    pipeline.EndADeclination = PipeEndA[3]

    # assign initial end coordinates for Winch

    pipelineWinch.ConnectionX[0] = WinchB[0]
    pipelineWinch.ConnectionZ[0] = WinchB[1]

    pipelineWinchX = pipelineWinch.ConnectionX[0]
    pipelineWinchTension = pipelineWinch.StageValue[0]

    model.SaveData(tarfilename)
    # extract tension and bending moment results
    pipeline_tension, pipeline_bending_moment, pipelineWinchTension = PEMS_Loads(tarfilename,pipelineWinchX)
    

#    print("\n")
    print("Initial Tension= ", round(pipeline_tension, 1), "Initial BM= ", abs(round(pipeline_bending_moment, 2)),
          "Initial WinchX=  ", round(pipelineWinchX, 3))

    # define target values
    target_tension = Tloads[0]
    target_bending_moment = Tloads[1]

    print("Targer Tension= ", round(target_tension, 1), "Target BM= ", abs(round(target_bending_moment, 2)), "\n")

    winchxi = pipelineWinchX
    # Check the moment to find the X-step direction
    func = lambda winchx: (PEMS_BM(tarfilename, winchx,target_tension)+target_bending_moment)
    
    winchxf = fsolve(func, winchxi-0.5,xtol=Tol)
    pipelineWinch.ConnectionX[0] = winchxf
    pipeline_tension_f, pipeline_bending_moment_f, pipelineWinchTension_f = PEMS_Loads(tarfilename, winchxf)                         
    
    pipelineWinch.StageValue[0] = pipelineWinchTension_f
    
    print("Final Tension= ", round(pipeline_tension_f, 3), "Final BM= ", abs(round(pipeline_bending_moment_f, 3)), "\n") 
    
    model.CalculateStatics()

    model.SaveData(tarfilename)

    RangeX = pipeline.RangeGraph('X')
    RangeZ = pipeline.RangeGraph('Z')
    #    print(RangeX.Max)
    return RangeX, RangeZ

    # define main body of code to extract data from spreadsheet


def PEMS_Analysis(filepathinfo):
    #    wb = xw.Book.caller()
    t0 = time.time()
    os.chdir(filepathinfo[0])
    wb = xw.Book(filepathinfo[0] + "\\" + filepathinfo[1])
    sh1 = wb.sheets('PEMS')
    sh1.range('status').value = 'EXTRACTING'
    
    shtemp = wb.sheets('temp')
    shhom = wb.sheets('HOM Exit Point')
    shten = wb.sheets('Tensioner Exit Point')
        
    HOM_Reel_pt = np.array(shhom.range('C8').options(expand='table').value)
    HOM_JL_pt = np.array(shhom.range('J8').options(expand='table').value)
    TEN_Reel_pt = np.array(shten.range('C8').options(expand='table').value)

    TEN_set_E = np.array(sh1.range('D15').options(expand='right').value)
    TEN_set_F = np.array(sh1.range('D16').options(expand='right').value)
    HOM_set_E = np.array(sh1.range('D17').options(expand='right').value)
    HOM_set_F = np.array(sh1.range('D18').options(expand='right').value)

    Eq_Set = [HOM_set_E, HOM_set_F, TEN_set_E, TEN_set_F]

    TEN_ten_E = sh1.range('H15').value
    TEN_ten_F = sh1.range('H16').value
    HOM_ten_E = sh1.range('H17').value
    HOM_ten_F = sh1.range('H18').value

    Eq_ten = [HOM_ten_E, HOM_ten_F, TEN_ten_E, TEN_ten_F]

    pipe_ID = sh1.range('pipe_ID').value

    Tol = sh1.range('tolerance').value

    TowerAngles = np.asarray(sh1.range('towerangles').options(expand='down').value, )
    HOMorTEN = sh1.range('HomorTen').options(expand='down').value
    EMPoFLD = sh1.range('emporfld').options(expand='down').value
    BatchRunData = [TowerAngles, HOMorTEN, EMPoFLD]

    filePath = filepathinfo[0]

    batchfile = open('csv1.bat', 'w')
    batchfile.write('copy ')

    sh1.range('status').value = 'RUNNING'
    for twn in range(np.size(TowerAngles)):

        if np.size(TowerAngles) == 1:
            twa = TowerAngles
            ht = HOMorTEN
            con = EMPoFLD
        else:
            twa = TowerAngles[twn]
            ht = HOMorTEN[twn]
            con = EMPoFLD[twn]
        if ht == 'HOM + TEN':
            ht = ['HOM', 'TEN']
        else:
            ht = np.array([ht])

        if con == 'Empty + Flooded':
            con = ['Empty', 'Flooded']
        else:
            con = np.array([con])

        for htn in ht:
            for cn in con:

                tenmax = eval(htn + '_ten_' + cn[0])
                tens = np.linspace(35, tenmax, num=4)

                LCFolder = str(twa) + '_' + htn + '_' + cn
                directory = filePath + "\\" + LCFolder
                #                directory = LCFolder
                if not os.path.exists(directory):
                    os.makedirs(directory)

                # pipe end position
                htrl = eval(htn + '_Reel_pt')
                pipeEndAX, pipeEndAZ = htrl[np.where(HOM_Reel_pt[:, 0] == twa)][0, 1:]
                pipeEndAY = 0
                #                pipeEndAZ = htrl[np.where(HOM_Reel_pt[:,0] == twa)][0,2]
                pipeEndAD = twa + 90

                pipeEndA = [pipeEndAX, pipeEndAY, pipeEndAZ, pipeEndAD]

                # winch end coordinates
                WinchEndBX = (0 - (-pipeEndAX)) - ((50 + 20) * np.cos(np.deg2rad(twa)))
                WinchEndBY = 0
                WinchEndBZ = (10.8 - (-pipeEndAZ)) - ((50 + 20) * np.sin(np.deg2rad(twa)))
                WinchB = [WinchEndBX, WinchEndBZ]

                shtemp.range('A1').value = (pipe_ID + "_" + htn + "_" + cn[0] + "_" + str(np.int(twa)))

                k = 0
                while k < 5:

                    # tensions and bending moment calculations
                    if k == 0:
                        tension = tenmax * 9.81
                        bendingmoment = 0
                    else:
                        tension = tens[k - 1] * 9.81
                        eqs = eval(htn + '_set_' + cn[0])
                        bendingmoment = eqs[0] * (tens[k - 1]) ** 2 + eqs[1] * (tens[k - 1]) + eqs[2]

                    FNameID = pipe_ID + '_' + str(twa) + 'deg_' + htn + '_' + cn + '_' + str(
                        np.int(tension)) + 'kN_' + str(np.int(bendingmoment)) + 'BM.dat'

                    FName1 = str(filePath) + "\\" + str(LCFolder) + "\\" + str(FNameID)
                    FName2 = str(filePath) + "\\" + str(LCFolder)
                    FileinfoData = [filePath, FName1]
                    TargetTen = tension
                    TargetBM = bendingmoment
                    Tloads = [TargetTen, TargetBM]
                    print('\n\n',twa, htn, cn)
                    [RangeX, RangeZ] = PEMS(FileinfoData, Tloads, pipeEndA, WinchB, cn, Tol)
                    #
                    #                    RangeL = RangeX.X

                    if k == 0:
                        RangeXstack = RangeX.Mean
                        RangeZstack = RangeZ.Mean
                    else:
                        RangeXstack = np.vstack((RangeXstack, RangeX.Mean))
                        RangeZstack = np.vstack((RangeZstack, RangeZ.Mean))

                    if k > 0:
                        shtemp.range((4, 3 * k - 2)).value = TargetTen
                        shtemp.range((4, 3 * k - 1)).value = TargetBM
                        RangeDeflection = np.sqrt(
                            (RangeXstack[k] - RangeXstack[0]) ** 2 + (RangeZstack[k] - RangeZstack[0]) ** 2)
                        shtemp.range((6, 3 * k - 1)).options(transpose=True).value = RangeDeflection

                    k += 1

                # shtemp.activate()
                rowsno = 137
                #
                writefile = open(LCFolder + '.csvtemp', 'w', newline='')
                file_writer = csv.writer(writefile, delimiter=',')
                #
                for r in range(rowsno):
                    tempr = shtemp.range((r + 1, 1), (r + 1, 11)).value
                    file_writer.writerow(tempr)

                writefile.close()
                batchfile.write(LCFolder + '.csvtemp ')
                batchfile.write('+')

    batchfile.close()

    batchfile = open('csv1.bat', 'rb+')
    batchfile.seek(-1, 2)
    batchfile.truncate()
    batchfile.close()

    batchfile = open('csv1.bat', 'a')
    batchfile.write(pipe_ID + '.csv\n')
    batchfile.write('del /s /q /f *.csvtemp')
    batchfile.close()

    p = Popen("csv1.bat")
    stdout, stderr = p.communicate()

    sh1.range('status').value = 'READY'
    os.remove('csv1.bat')
    
    t1 = time.time()
    print(t1 - t0)


if __name__ == "__main__":
    try:
        filePath = sys.argv[1]
        wbName = sys.argv[2]

    except IndexError:
        print("python script NOT found.")
        filePath = (sys.argv[0]).strip('/PEMS.py')
        wbName = 'PEMS.xlsm'

    filepathinfo = [filePath, wbName]
    PEMS_Analysis(filepathinfo)
    os.remove('TEMP0.bat')	