#%% Script data

# Ehsan Alizadeh
# LinkedIn: https://www.linkedin.com/in/ehsanalizd

#%% Input

file_path = r"D:\Thesis"

#%% Import Libraries

import os
import sys
import comtypes.client
import pandas as pd
from pyautocad import Autocad
import numpy as np
import json
import inspect
import math
from scipy.interpolate import interp1d

#%% Define Functions

file_path = file_path.replace("\\", "/")

def compute_arc_points(start, end, third, num_segments):
    start = np.array(start)
    end = np.array(end)
    third = np.array(third)
    D = 2 * (start[0] * (third[1] - end[1]) + third[0] * (end[1] - start[1]) + end[0] * (start[1] - third[1]))
    centerX = ((start[0]**2 + start[1]**2) * (third[1] - end[1]) +
               (third[0]**2 + third[1]**2) * (end[1] - start[1]) +
               (end[0]**2 + end[1]**2) * (start[1] - third[1])) / D
    centerY = ((start[0]**2 + start[1]**2) * (end[0] - third[0]) +
               (third[0]**2 + third[1]**2) * (start[0] - end[0]) +
               (end[0]**2 + end[1]**2) * (third[0] - start[0])) / D
    radius = np.sqrt((centerX - start[0])**2 + (centerY - start[1])**2)
    theta_start = np.arctan2(start[1] - centerY, start[0] - centerX)
    theta_end = np.arctan2(end[1] - centerY, end[0] - centerX)
    theta_third = np.arctan2(third[1] - centerY, third[0] - centerX)
    if theta_start < 0:
        theta_start += 2 * np.pi
    if theta_end < 0:
        theta_end += 2 * np.pi
    if theta_third < 0:
        theta_third += 2 * np.pi
    angles = np.linspace(theta_start, theta_end, num_segments + 1)
    arc_points = []
    for angle in angles:
        x = round(centerX + radius * np.cos(angle), 4)
        y = 0
        z = round(centerY + radius * np.sin(angle), 4)
        arc_points.append([x, y, z])
    return arc_points


def save_variables(filename, *args):
    # Load existing data
    if os.path.exists(filename):
        with open(filename, 'r', encoding='utf-8') as file:
            try:
                existing_variables = json.load(file)
            except json.JSONDecodeError:
                existing_variables = {}
    else:
        existing_variables = {}

    # Update or add new variables
    caller_locals = inspect.currentframe().f_back.f_locals
    for name in args:
        if name in caller_locals:
            value = caller_locals[name]
            
            # Convert NumPy arrays to lists
            if isinstance(value, np.ndarray):
                value = {"__type__": "ndarray", "data": value.tolist()}
            
            # Convert DataFrames to dictionaries with a marker
            elif isinstance(value, pd.DataFrame):
                value = {"__type__": "dataframe", "data": value.to_dict(orient="records")}

            existing_variables[name] = value

    # Save updated data
    with open(filename, 'w', encoding='utf-8') as file:
        json.dump(existing_variables, file, ensure_ascii=False, indent=4)

def load_variables(filename):
    if os.path.exists(filename):
        with open(filename, 'r', encoding='utf-8') as file:
            try:
                data = json.load(file)
            except json.JSONDecodeError:
                return {}

        # Restore DataFrames and NumPy arrays
        for key, value in data.items():
            if isinstance(value, dict) and "__type__" in value:
                if value["__type__"] == "ndarray":
                    data[key] = np.array(value["data"])  # Convert back to NumPy array
                elif value["__type__"] == "dataframe":
                    data[key] = pd.DataFrame(value["data"])  # Convert back to DataFrame
        return data

    return {}


def distance(p1, p2):
    return math.sqrt((p2[0] - p1[0])**2 + (p2[1] - p1[1])**2)


def arc_length(p1, p2, p3):
    a = distance(p1, p2)
    b = distance(p2, p3)
    c = distance(p1, p3)
    s = (a + b + c) / 2
    radius = (a * b * c) / (4 * math.sqrt(s * (s - a) * (s - b) * (s - c)))
    angle_rad = 2 * math.asin(c / (2 * radius))
    arc_length = radius * angle_rad
    return arc_length


def ImportSheet(Database, Sheet, Cols):
    df = pd.read_excel(Database, sheet_name=Sheet, usecols = Cols, header=[1])
    df = df.drop(0)
    df = df.reset_index(drop=True)
    return df

#%% Connect to ETABS

AttachToInstance = True
helper = comtypes.client.CreateObject('ETABSv1.Helper')
helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)

if AttachToInstance:
    #attach to a running instance of ETABS
    try:
        #get the active ETABS object
        myETABSObject = helper.GetObject("CSI.ETABS.API.ETABSObject") 
    except (OSError, comtypes.COMError):
        print("No running instance of the program found or failed to attach.")
        sys.exit(-1)

#create SapModel object
SapModel = myETABSObject.SapModel

#%% Draw Frames

N_mm_C = 9
SapModel.SetPresentUnits(N_mm_C)
h = 3500
b = 7000
bay = 4
brbay = [2]
straight_leng = (h**2 + (b/2)**2)**0.5     # leng: length of straight line of braces
kisi = "straight_leng/500"                        # e: the amount of eccentricity
e = eval(kisi)
BrTy='Bilinear'                   # Crescent or Bilinear
DrawAutoCAD = "No"
if DrawAutoCAD == "Yes" :
    # Connect to AutoCAD
    acad = Autocad(create_if_not_exists=True)

save_vars = "[h, b, bay, brbay, straight_leng, kisi, e, BrTy]"
save_vars = save_vars.strip('[]').split(',')
save_vars = [element.strip() for element in save_vars]
save_variables(f"{file_path}/Database.json", *save_vars)

#%% Draw Columns

loaded_variables = load_variables(f"{file_path}/Database.json")
bay = loaded_variables['bay']
b = loaded_variables['b']
h = loaded_variables['h']

ColNum= bay + 1
Cx1, Cx2, Cy1, Cy2, Cz1, Cz2 = [], [], [], [], [], []
for i in range(ColNum):
    Cx1new = float(i*b)
    Cx1.append(Cx1new)
    Cy1.append(0.0)
    Cy2.append(0.0)
    Cz1.append(0.0)
    Cz2.append(h)
Cx2 = Cx1[:]
for i in range(ColNum):
    # SapModel.FrameObj.AddByCoord(xi, yi, zi, xj, yj, zj, Name, PropName)
    SapModel.FrameObj.AddByCoord(Cx1[i], Cy1[i], Cz1[i], Cx2[i], Cy2[i], Cz2[i], "", "AutoColumn")
    if DrawAutoCAD == "Yes" :
        points=[(Cx1[i],Cz1[i]), (Cx2[i],Cz2[i])]
        points_str=" ".join(f"{x},{y}" for x, y in points)
        acad.doc.SendCommand(f"LINE {points_str} \n ")  

save_vars = "[ColNum, Cx1, Cx2, Cy1, Cy2, Cz1, Cz2, Cx1new]"
save_vars = save_vars.strip('[]').split(',')
save_vars = [element.strip() for element in save_vars]
save_variables(f"{file_path}/Database.json", *save_vars)

#%% Draw Beams

loaded_variables = load_variables(f"{file_path}/Database.json")
bay = loaded_variables['bay']
b = loaded_variables['b']
h = loaded_variables['h']

BeamNum = bay
Bx1, Bx2, By1, By2, Bz1, Bz2 = [], [], [], [], [], []
for i in range(BeamNum):
    Bx1new = float(i * b)
    Bx1.append(Bx1new)
for i in range(1, BeamNum+1):
    Bx2new = float(i * b)
    Bx2.append(Bx2new)
for _ in range(BeamNum+1):
    By1.append(0.0)
    By2.append(0.0)
    Bz1.append(h)
    Bz2.append(h)
for i in range(BeamNum):
    # SapModel.FrameObj.AddByCoord(xi, yi, zi, xj, yj, zj, Name, PropName)
    SapModel.FrameObj.AddByCoord(Bx1[i], By1[i], Bz1[i], Bx2[i], By2[i], Bz2[i], "", "AutoBeam")
    if DrawAutoCAD == "Yes" :
        points=[(Bx1[i], Bz1[i]), (Bx2[i], Bz2[i])]
        points_str=" ".join(f"{x},{y}" for x, y in points)
        acad.doc.SendCommand(f"LINE {points_str} \n ")

save_vars = "[BeamNum, Bx1, Bx2, By1, By2, Bz1, Bz2]"
save_vars = save_vars.strip('[]').split(',')
save_vars = [element.strip() for element in save_vars]
save_variables(f"{file_path}/Database.json", *save_vars)

#%% Draw Braces

loaded_variables = load_variables(f"{file_path}/Database.json")
bay = loaded_variables['bay']
b = loaded_variables['b']
h = loaded_variables['h']
brbay = loaded_variables['brbay']
e = loaded_variables['e']
Cx1 = loaded_variables['Cx1']
Cx2 = loaded_variables['Cx2']
Cy1 = loaded_variables['Cy1']
Cy2 = loaded_variables['Cy2']
Cz1 = loaded_variables['Cz1']
Cz2 = loaded_variables['Cz2']
Bx1 = loaded_variables['Bx1']
Bx2 = loaded_variables['Bx2']
By1 = loaded_variables['By1']
By2 = loaded_variables['By2']
Bz1 = loaded_variables['Bz1']
Bz2 = loaded_variables['Bz2']

BraceNum = len(brbay)*2
brpoint = [x-1 for x in brbay]
Brconx = [] # Brconx: Braces connection point (x-coordinate)
for i in range(bay):
    Brconxnew = (Bx2[i]+Bx1[i])/2
    Brconx.append(Brconxnew)
Brcony = [] # Brconx: Braces connection point (y-coordinate)
for _ in range(bay):
    Brcony.append(0.0)
Brconz = []
for _ in range(bay):
    Brconz.append(h)
    Brmidxl = [] # Brmidxl: mid point of left brace (x-coordinate)
Brmidyl = []
Brmidzl = []
for i in range(bay):
    Brmidxlnew = (Cx1[i] + Brconx[i])/2
    Brmidxl.append(Brmidxlnew)
    Brmidyl.append(0.0)
    Brmidzlnew = (Cz1[i] + Brconz[i])/2
    Brmidzl.append(Brmidzlnew)
Brmidxr = [] # Brmidxr: mid point of right brace (x-coordinate)
Brmidyr = []
Brmidzr = []
for i in range(bay):
    Brmidxrnew = (Cx1[i+1]+Brconx[i])/2
    Brmidxr.append(Brmidxrnew)
    Brmidyr.append(0.0)
    Brmidzrnew = (Cz1[i+1]+Brconz[i])/2
    Brmidzr.append(Brmidzrnew)

ACleng = []
for i in range(bay):
    AClengnew = (((Cx1[i]-Brconx[i])**2 + (Cz1[i]-Brconz[i])**2)**0.5)/2
    ACleng.append(AClengnew)

Brcurxl, Brcuryl, Brcurzl = [], [], []
for i in range(bay):
    Brcurxlnew = Brmidxl[i] - (e*(Brmidzl[i]-Cz1[i]))/ACleng[i]
    Brcurxl.append(Brcurxlnew)
    Brcurzlnew = Brmidzl[i] + (e*(Brmidxl[i]-Cx1[i]))/ACleng[i]
    Brcurzl.append(Brcurzlnew)
for _ in range(bay):
    Brcuryl.append(0.0)
Brcurxr, Brcuryr, Brcurzr = [], [], []
for i in range(bay):
    Brcurxrnew = Brmidxr[i] + (e*(Brmidzr[i]-Cz1[i+1]))/ACleng[i]
    Brcurxr.append(Brcurxrnew)
    Brcurzrnew = Brmidzr[i] - (e*(Brmidxr[i]-Cx1[i+1]))/ACleng[i]
    Brcurzr.append(Brcurzrnew)

for _ in range(bay):
    Brcuryr.append(0.0)
if BrTy=='Crescent' :
    SegNum = 14
elif BrTy=='Bilinear' :
    SegNum = 2

for i in brpoint:
    start_left = (Cx1[i], Cz1[i])
    end_left = (Brconx[i], Brconz[i])
    mid_left = (Brcurxl[i], Brcurzl[i])
    points_left = compute_arc_points(start_left, end_left, mid_left, SegNum)
    
    start_right = (Cx1[i+1], Cz1[i+1])
    end_right = (Brconx[i], Brconz[i])
    mid_right = (Brcurxr[i], Brcurzr[i])
    points_right = compute_arc_points(start_right, end_right, mid_right, SegNum)

    for j in range(SegNum):
        # SapModel.FrameObj.AddByCoord(xi, yi, zi, xj, yj, zj, Name, PropName)
        SapModel.FrameObj.AddByCoord(*points_left[j], *points_left[j+1], "", "AutoBrace")
    
    for j in range(SegNum):
        # SapModel.FrameObj.AddByCoord(xi, yi, zi, xj, yj, zj, Name, PropName)
        SapModel.FrameObj.AddByCoord(*points_right[j], *points_right[j+1], "", "AutoBrace")

i = brpoint[0]
ArcLen = arc_length((Cx1[i], Cz1[i]), (Brcurxl[i], Brcurzl[i]), (Brconx[i], Brconz[i]))

save_vars = "[BraceNum, ArcLen, SegNum]"
save_vars = save_vars.strip('[]').split(',')
save_vars = [element.strip() for element in save_vars]
save_variables(f"{file_path}/Database.json", *save_vars)

#%% Assign End Releases

loaded_variables = load_variables(f"{file_path}/Database.json")
bay = loaded_variables['bay']
ColNum = loaded_variables['ColNum']
BeamNum = loaded_variables['BeamNum']
BraceNum = loaded_variables['BraceNum']
SegNum = loaded_variables['SegNum']

Beam_EndRelease = "Yes"
blab = []
first_beam = bay+2
blab.append(first_beam)
for j in range(1, bay):
    blab.append(first_beam+j)

if Beam_EndRelease=="Yes" :
    ii = [False] * 6
    jj = [False] * 6
    StartValue = [0.0] * 6
    EndValue = [0.0] * 6
    ii[5] = True
    jj[5] = True
    for i in blab:
        # SapModel.FrameObj.SetReleases(Name, ii, jj, StartValue, EndValue)
        SapModel.FrameObj.SetReleases(str(i), ii, jj, StartValue, EndValue)

BrTag1List = []
BrTag2List = []
for q in range(BraceNum):
    tag1 = ColNum+BeamNum+(q*SegNum)+1
    tag2 = ColNum+BeamNum+(q*SegNum)+SegNum
    BrTag1List.append(tag1)
    BrTag2List.append(tag2)

Brace_EndRelease = "Yes"
if Brace_EndRelease=="Yes" :
    ii = [False] * 6
    jj = [False] * 6
    StartValue = [0.0] * 6
    EndValue = [0.0] * 6
    ii[5] = True
    for i in BrTag1List:
        # SapModel.FrameObj.SetReleases(Name, ii, jj, StartValue, EndValue)
        SapModel.FrameObj.SetReleases(str(i), ii, jj, StartValue, EndValue)
    ii = [False] * 6
    jj = [False] * 6
    StartValue = [0.0] * 6
    EndValue = [0.0] * 6
    jj[5] = True
    for i in BrTag2List:
        # SapModel.FrameObj.SetReleases(Name, ii, jj, StartValue, EndValue)
        SapModel.FrameObj.SetReleases(str(i), ii, jj, StartValue, EndValue)

#%% Assign Joint Restraints

loaded_variables = load_variables(f"{file_path}/Database.json")
ColNum = loaded_variables['ColNum']

ColumnRestraints = "Pinned"
if ColumnRestraints == "Fixed" :
    for i in range(ColNum):
        # Restraint = [U1, U2, U3, R1, R2, R3]
        Restraint = [True, True, True, True, True, True]
        # SapModel.PointObj.SetRestraint(Name, Value)
        SapModel.PointObj.SetRestraint(str(2*i+1), Restraint)
if ColumnRestraints == "Pinned" :
    for i in range(ColNum):
        # Restraint = [U1, U2, U3, R1, R2, R3]
        Restraint = [True, True, True, False, False, False]
        # SapModel.PointObj.SetRestraint(Name, Value)
        SapModel.PointObj.SetRestraint(str(2*i+1), Restraint)

#%% Leaning Column

loaded_variables = load_variables(f"{file_path}/Database.json")
Cx1new = loaded_variables['Cx1new']
b = loaded_variables['b']
h = loaded_variables['h']

LCx1 = Cx1new+b/2
LCy1 = 0
LCz1 = 0
LCx2 = Cx1new+b/2
LCy2 = 0
LCz2 = h
# SapModel.FrameObj.AddByCoord(xi, yi, zi, xj, yj, zj, Name, PropName)
[leanCollabel, ret] = SapModel.FrameObj.AddByCoord(LCx1, LCy1, LCz1, LCx2, LCy2, LCz2, "", "W1100X607")
# LeanColPropModif = [Area, Shear2, Shear3, Torsional, Moment2, Moment3, Mass, Weight]
LeanColPropModif = [10000000.0, 1.0, 1.0, 1.0, 1.0, 1.0, 0.0, 0.0]
# SapModel.PropFrame.SetModifiers(name, value)
SapModel.FrameObj.SetModifiers(leanCollabel, LeanColPropModif)
ii = [False] * 6
jj = [False] * 6
StartValue = [0.0] * 6
EndValue = [0.0] * 6
ii[5] = True
# SapModel.FrameObj.SetReleases(Name, ii, jj, StartValue, EndValue)
SapModel.FrameObj.SetReleases(leanCollabel, ii, jj, StartValue, EndValue)

#%% Leaning Beam

loaded_variables = load_variables(f"{file_path}/Database.json")
Cx1new = loaded_variables['Cx1new']
b = loaded_variables['b']
h = loaded_variables['h']

LBx1 = Cx1new
LBy1 = 0
LBz1 = h
LBx2 = Cx1new+b/2
LBy2 = 0
LBz2 = h
# SapModel.FrameObj.AddByCoord(xi, yi, zi, xj, yj, zj, Name, PropName)
[leanBeamlabel, ret] = SapModel.FrameObj.AddByCoord(LBx1, LBy1, LBz1, LBx2, LBy2, LBz2, "", "W100X19.3")
ii = [False] * 6
jj = [False] * 6
StartValue = [0.0] * 6
EndValue = [0.0] * 6
ii[5] = True
jj[5] = True
# SapModel.FrameObj.SetReleases(Name, ii, jj, StartValue, EndValue)
SapModel.FrameObj.SetReleases(leanBeamlabel, ii, jj, StartValue, EndValue)
# LeanBeamPropModif = [Area, Shear2, Shear3, Torsional, Moment2, Moment3, Mass, Weight]
LeanBeamPropModif = [10000000.0, 1.0, 1.0, 1.0, 1.0, 1.0, 0.0, 0.0]
# SapModel.PropFrame.SetModifiers(name, value)
SapModel.FrameObj.SetModifiers(leanBeamlabel, LeanBeamPropModif)

#%% Assign Frame Gravity Loads

loaded_variables = load_variables(f"{file_path}/Database.json")
b = loaded_variables['b']
h = loaded_variables['h']

SuperDeadShell = 0.003
SlabShell = 0.0025
PartitionShell = 0.001
LiveShell = 0.002

SuperDeadLoad = SuperDeadShell * b/2
SlabLoad = SlabShell * b/2
PartitionLoad = PartitionShell * b/2
WallLoad = 8.0
LiveLoad = LiveShell * b/2

# SapModel.FrameObj.SetLoadDistributed(Name, LoadPat, MyType, Dir, Dist1, Dist2, Val1, Val2, CSys, RelDist, Replace, ItemType)
SapModel.FrameObj.SetLoadDistributed("Beams (ALL)", "SuperDead", 1, 10, 0, 1, SuperDeadLoad, SuperDeadLoad, "Global", True, True, 1)
SapModel.FrameObj.SetLoadDistributed("Beams (ALL)", "Slab", 1, 10, 0, 1, SlabLoad, SlabLoad, "Global", True, True, 1)
SapModel.FrameObj.SetLoadDistributed("Beams (ALL)", "Partition", 1, 10, 0, 1, PartitionLoad, PartitionLoad, "Global", True, True, 1)
SapModel.FrameObj.SetLoadDistributed("Beams (ALL)", "Wall", 1, 10, 0, 1, WallLoad, WallLoad, "Global", True, True, 1)
SapModel.FrameObj.SetLoadDistributed("Beams (ALL)", "Live", 1, 10, 0, 1, LiveLoad, LiveLoad, "Global", True, True, 1)

WallPointLoad = b/2 * WallLoad
# SapModel.PointObj.SetLoadForce(Name, LoadPat, Value, Replace, CSys, ItemType)
Value = [0, 0, -WallPointLoad, 0, 0, 0]
SapModel.PointObj.SetLoadForce("Ecl Fixed Nodes", "Wall", Value, True, "Global", 1)

save_vars = "[SuperDeadShell, SlabShell, PartitionShell, LiveShell, SuperDeadLoad, SlabLoad, PartitionLoad, WallLoad, LiveLoad, WallPointLoad]"
save_vars = save_vars.strip('[]').split(',')
save_vars = [element.strip() for element in save_vars]
save_variables(f"{file_path}/Database.json", *save_vars)

#%% Leaning Type1 Loads

loaded_variables = load_variables(f"{file_path}/Database.json")
b = loaded_variables['b']
bay = loaded_variables['bay']
SuperDeadShell = loaded_variables['SuperDeadShell']
SlabShell = loaded_variables['SlabShell']
PartitionShell = loaded_variables['PartitionShell']
WallLoad = loaded_variables['WallLoad']
LiveShell = loaded_variables['LiveShell']

SuperDeadLoadLeaning1 = SuperDeadShell * 3*b/2 * (b*bay)
SlabLoadLeaning1 = SlabShell * 3*b/2 * (b*bay)
PartitionLoadLeaning1 = PartitionShell * 3*b/2 * (b*bay)
WallLoadLeaning1 = 2 * (3*b/2) * WallLoad
LiveLoadLeaning1 = LiveShell * 3*b/2 * (b*bay)

LeaningDeads1 = SuperDeadLoadLeaning1 + SlabLoadLeaning1 + PartitionLoadLeaning1 + WallLoadLeaning1
LeaningLives1 = LiveLoadLeaning1

# SapModel.PointObj.SetLoadForce(Name, LoadPat, Value, Replace, CSys, ItemType)
ValueDeads1 = [0, 0, -LeaningDeads1, 0, 0, 0]
SapModel.PointObj.SetLoadForce("Leaning Joint", "L1D", ValueDeads1, True, "Global", 1)
ValueLives1 = [0, 0, -LeaningLives1, 0, 0, 0]
SapModel.PointObj.SetLoadForce("Leaning Joint", "L1L", ValueLives1, True, "Global", 1)

save_vars = "[SuperDeadLoadLeaning1, SlabLoadLeaning1, PartitionLoadLeaning1, WallLoadLeaning1, LiveLoadLeaning1, LeaningDeads1, LeaningLives1]"
save_vars = save_vars.strip('[]').split(',')
save_vars = [element.strip() for element in save_vars]
save_variables(f"{file_path}/Database.json", *save_vars)

#%% Leaning Type2 Loads

loaded_variables = load_variables(f"{file_path}/Database.json")
b = loaded_variables['b']
SuperDeadShell = loaded_variables['SuperDeadShell']
SlabShell = loaded_variables['SlabShell']
PartitionShell = loaded_variables['PartitionShell']
WallLoad = loaded_variables['WallLoad']
LiveShell = loaded_variables['LiveShell']

SuperDeadLoadLeaning2 = SuperDeadShell * 3*b/2
SlabLoadLeaning2 = SlabShell * 3*b/2
PartitionLoadLeaning2 = PartitionShell * 3*b/2
WallLoadLeaning2 = WallLoad
LiveLoadLeaning2 = LiveShell * 3*b/2

LeaningDeads2 = SuperDeadLoadLeaning2 + SlabLoadLeaning2 + PartitionLoadLeaning2 + WallLoadLeaning2
LeaningLives2 = LiveLoadLeaning2

# SapModel.FrameObj.SetLoadDistributed(Name, LoadPat, MyType, Dir, Dist1, Dist2, Val1, Val2, CSys, RelDist, Replace, ItemType)
SapModel.FrameObj.SetLoadDistributed("Beams (ALL)", "L2D", 1, 10, 0, 1, LeaningDeads2, LeaningDeads2, "Global", True, True, 1)
SapModel.FrameObj.SetLoadDistributed("Beams (ALL)", "L2L", 1, 10, 0, 1, LeaningLives2, LeaningLives2, "Global", True, True, 1)
WallLoadLeaning22 = (3/2*b) * WallLoad
Value = [0, 0, -WallLoadLeaning22, 0, 0, 0]
SapModel.PointObj.SetLoadForce("Ecl Fixed Nodes", "L2D", Value, True, "Global", 1)

save_vars = "[SuperDeadLoadLeaning2, SlabLoadLeaning2, PartitionLoadLeaning2, WallLoadLeaning2, LiveLoadLeaning2, LeaningDeads2, LeaningLives2, WallLoadLeaning22]"
save_vars = save_vars.strip('[]').split(',')
save_vars = [element.strip() for element in save_vars]
save_variables(f"{file_path}/Database.json", *save_vars)

#%% Set Diaphragm

GroupAssigns = SapModel.GroupDef.GetAssignments("Panel Zone")
GroupAssigns = GroupAssigns[2]

for i in range(len(GroupAssigns)):
    SapModel.PointObj.SetDiaphragm(str(GroupAssigns[i]), 3, "D1")

#%% Define Load Combinations

SapModel.DesignSteel.SetCode("AISC 360-22")
# Framing Type: SCBF
SapModel.DesignSteel.AISC360_22.SetPreference(3, 4)
# S_ds for vertical earthquake load
S_DS = 1.0
SapModel.DesignSteel.AISC360_22.SetPreference(7, S_DS)
# Add notional loads
SapModel.DesignSteel.AISC360_22.SetPreference(15, 1)
# Demand/Capacity Ratio
SapModel.DesignSteel.AISC360_22.SetPreference(37, 1)

# SapModel.RespCombo.AddDesignDefaultCombos(DesignSteel, DesignConcrete, DesignAluminum, DesignColdFormed)
SapModel.RespCombo.AddDesignDefaultCombos(True, False, False, False)

#%% Assign Unbraced Length Ratios

# SapModel.DesignSteel.AISC360_22.SetOverwrite(Name, Item, Value, ItemType)
SapModel.DesignSteel.AISC360_22.SetOverwrite("Beams (ALL)", 26, 0.01, 1)
SapModel.DesignSteel.AISC360_22.SetOverwrite("Beams (ALL)", 27, 0.01, 1)
SapModel.DesignSteel.AISC360_22.SetOverwrite("Beams (ALL)", 28, 0.01, 1)

#%% Set Active degrees of freedom

DOF = [True, False, True, False, True, False]
SapModel.Analyze.SetActiveDOF(DOF)

#%% Insertion Point 

SapModel.FrameObj.SetInsertionPoint("Beams (ALL)", 8, False, False, [0, 0, 0], [0, 0, 0], "Global", 1)

#%% Define Lateral Seismic Load

SiteClass = "D"
RiskCategory = "II"
I_e = 1.0
R = 6.0
S_S = 1.5
S_1 = 0.6
F_a = 1.0
F_v = 1.5

S_MS = F_a * S_S
S_M1 = F_v * S_1
S_DS = (2/3) * S_MS
S_D1 = (2/3) * S_M1

if S_DS<0.167 and (RiskCategory == "I" or RiskCategory == "II" or RiskCategory == "III" or RiskCategory == "IV"):
    SDC1 = "A"
if S_DS>=0.167 and S_DS<0.33 and (RiskCategory == "I" or RiskCategory == "II" or RiskCategory == "III"):
    SDC1 = "B"
if S_DS>=0.167 and S_DS<0.33 and (RiskCategory == "IV"):
    SDC1 = "C"
if S_DS>=0.33 and S_DS<0.5 and (RiskCategory == "I" or RiskCategory == "II" or RiskCategory == "III"):
    SDC1 = "C"
if S_DS>=0.33 and S_DS<0.5 and (RiskCategory == "IV"):
    SDC1 = "D"
if S_DS>=0.5 and (RiskCategory == "I" or RiskCategory == "II" or RiskCategory == "III" or RiskCategory == "IV"):
    SDC1 = "D"  

if S_D1<0.067 and (RiskCategory == "I" or RiskCategory == "II" or RiskCategory == "III" or RiskCategory == "IV"):
    SDC2 = "A"
if S_D1>=0.067 and S_D1<0.133 and (RiskCategory == "I" or RiskCategory == "II" or RiskCategory == "III"):
    SDC2 = "B"
if S_D1>=0.067 and S_D1<0.133 and (RiskCategory == "IV"):
    SDC2 = "C"
if S_D1>=0.133 and S_D1<0.2 and (RiskCategory == "I" or RiskCategory == "II" or RiskCategory == "III"):
    SDC2 = "C"
if S_D1>=0.133 and S_D1<0.2 and (RiskCategory == "IV"):
    SDC2 = "D"
if S_D1>=0.2 and (RiskCategory == "I" or RiskCategory == "II" or RiskCategory == "III"):
    SDC2 = "D"  
if S_D1>=0.2 and (RiskCategory == "IV"):
    SDC2 = "N.A."  
SDC = max((SDC1, SDC2), key=ord)

N_m_C = 10
SapModel.SetPresentUnits(N_m_C)
[BaseElevation, NumberStories, StoryNam, StoryElevations, StoryHeights, IsMasterStory, SimilarToStory, SpliceAbove, SpliceHeight, color, ret] = SapModel.Story.GetStories_2()
h_n = round(max(StoryElevations), 1)
x = 0.75
C_t = 0.0488
T_A = C_t * h_n**x
T_Mode1_ETABS = 10
sd1_values = [0.4, 0.3, 0.2, 0.15, 0.1]
cu_values = [1.4, 1.4, 1.5, 1.6, 1.7]
interpolation_function = interp1d(sd1_values, cu_values, kind='linear', fill_value='extrapolate')
C_u = interpolation_function(S_D1)
T_Fun = C_u * T_A
T = min(T_Fun, T_Mode1_ETABS)
T_0 = 0.2*(S_D1/S_DS)
T_S = (S_D1/S_DS)
T_L = 6
if T <= T_0:
    S_a = S_DS*(0.4+0.6*(T/T_0))
elif T > T_0 and T <= T_S:
    S_a = S_DS
elif T > T_S and T <= T_L:
    S_a = S_D1/T
elif T>= T_L:
    S_a = S_D1*T_L/T**2

C_cal = S_a/(R/I_e)

if T <= T_L :
    C_max = S_D1/T/(R/I_e)
if T > T_L :
    C_max = S_D1*T_L/T**2/(R/I_e)

if S_1 < 0.6:
    C_min = max(0.044*S_DS*I_e, 0.01)
if S_1 >= 0.6:
    C_min = 0.5*S_1/(R/I_e)

if C_cal <= C_min:
    C_s = round(C_min, 4)
elif C_cal >= C_max:
    C_s = round(C_max, 4)
else:
    C_s = round(C_cal, 4)

if T < 0.5:
    K = 1.0 
elif T > 2.5:
    K = 2.0
else:
    K = float((T - 0.5) / (2.5 - 0.5) * (2 - 1) + 1)

print(f"SDC: {SDC}")
print(f"Cs: {round(C_s, 3)}")
print(f"K: {round(K, 3)}")

save_vars = "[I_e, R, S_S, S_1, F_a, F_v, S_MS, S_M1, S_DS, S_D1, SDC, h_n, x, C_t, T_A, C_u, T_Fun, T, T_0, T_S, T_L, S_a, C_cal, C_max, C_min, , C_s, K]"
save_vars = save_vars.strip('[]').split(',')
save_vars = [element.strip() for element in save_vars]
save_variables(f"{file_path}/Database.json", *save_vars)

#%% Assign Ecl point loads

loaded_variables = load_variables(f"{file_path}/Database.json")
eFy_A500GrB = loaded_variables['eFy_A500GrB']
ArcLen = loaded_variables['ArcLen']
h = loaded_variables['h']
b = loaded_variables['b']
e = loaded_variables['e']
straight_leng = loaded_variables['straight_leng']

RealArcLength = ArcLen

GetAllFrames = SapModel.FrameObj.GetAllFrames()
FrameList = GetAllFrames[1]

# from scipy.optimize import fsolve
# def radius_equation(r, arc_length, chord_length):
#     theta = arc_length / r
#     return chord_length - (2 * r * math.sin(theta / 2))

# def calculate_radius(arc_length, chord_length):
#     initial_guess = chord_length / 2
#     radius_solution = fsolve(radius_equation, initial_guess, args=(arc_length, chord_length))
#     return radius_solution[0]

# radius = calculate_radius(ArcLen, straight_leng)

# tethaParameter = RealArcLength/radius
# e = radius*(1-math.cos(tethaParameter/2))

N_mm_C = 9
SapModel.SetPresentUnits(N_mm_C)
Section_Properties = r"D:\Thesis\3. Codes\Section Properties.xlsx"
Frame_Prop = ImportSheet(Section_Properties, "Frame Prop - Summary", ["Name", "Material", "Shape", "Area", "R33", "Z33"])
Frame_Prop = Frame_Prop.drop(0)
Frame_Prop = Frame_Prop.drop(1)
Frame_Prop = Frame_Prop.drop(2)
Frame_Prop = Frame_Prop.reset_index(drop = True)
FrameSec_steelTube = ImportSheet(Section_Properties, "Frame Sec Def - Steel Tube", ["Name", "Material", "Total Depth", "Total Width", "Flange Thickness", "Web Thickness"])
FrameSec_steelTube = FrameSec_steelTube.merge(Frame_Prop, on='Name', how='left', suffixes=('', '_big'))
FrameSec_steelTube = FrameSec_steelTube.drop(['Material_big', 'Shape'], axis=1)
FrameSec_steelTube = FrameSec_steelTube.set_index("Name")

[NumberItems, ObjectType, ObjectName, ret] = SapModel.GroupDef.GetAssignments("Braces (End I-L)")
EndIleft = pd.DataFrame({'ObjectName': ObjectName})
EndIleft['BrLoc'] = 'End I-L'

[NumberItems, ObjectType, ObjectName, ret] = SapModel.GroupDef.GetAssignments("Braces (End J-L)")
EndJleft = pd.DataFrame({'ObjectName': ObjectName})
EndJleft['BrLoc'] = 'End J-L'

[NumberItems, ObjectType, ObjectName, ret] = SapModel.GroupDef.GetAssignments("Braces (End I-R)")
EndIright = pd.DataFrame({'ObjectName': ObjectName})
EndIright['BrLoc'] = 'End I-R'

[NumberItems, ObjectType, ObjectName, ret] = SapModel.GroupDef.GetAssignments("Braces (End J-R)")
EndJright = pd.DataFrame({'ObjectName': ObjectName})
EndJright['BrLoc'] = 'End J-R'

BrEcl = pd.concat([EndIleft, EndJleft, EndIright, EndJright], ignore_index=True)

BrEcl['PutoFiPn'] = 0
BrEcl['PutoFiPn'] = BrEcl['PutoFiPn'].astype(float)
BrEcl['Ecl_Tn'] = 0
BrEcl['Ecl_Tn'] = BrEcl['PutoFiPn'].astype(float)
BrEcl['Ecl_Pn'] = 0
BrEcl['Ecl_Pn'] = BrEcl['PutoFiPn'].astype(float)
BrEcl['Ecl_03Pn'] = 0
BrEcl['Ecl_03Pn'] = BrEcl['PutoFiPn'].astype(float)
BrEcl['ObjectName'] = BrEcl['ObjectName'].astype(int)
BrEcl = BrEcl.sort_values(by='ObjectName')
BrEcl = BrEcl.reset_index(drop=True)

# for i in range(len(BrEcl)):
#     if BrEcl.loc[i, 'BrLoc'] == "End I-L" or BrEcl.loc[i, 'BrLoc'] == "End I-R":
#         BrEcl.loc[i, 'ObjectName'] = BrEcl.loc[i, 'ObjectName']+4
#     if BrEcl.loc[i, 'BrLoc'] == "End J-L" or BrEcl.loc[i, 'BrLoc'] == "End J-R":
#         BrEcl.loc[i, 'ObjectName'] = BrEcl.loc[i, 'ObjectName']-4

for i in [x for x in range(len(BrEcl)) if x % 2 == 0] : 
    AxialForceList = []
    for u in range(BrEcl.loc[i, 'ObjectName'], BrEcl.loc[i+1, 'ObjectName']+1):
        DesignSteelResult1 = SapModel.DesignResults.DesignForces.ColumnDesignForces(str(u))
        try:
            MaxDesignSteelResult = max([abs(t) for t in DesignSteelResult1[4]])
        except:
            pass
        
        DesignSteelResult2 = SapModel.DesignResults.DesignForces.BeamDesignForces(str(u))
        try:
            MaxDesignSteelResult = max([abs(t) for t in DesignSteelResult2[4]])
        except:
            pass
        
        DesignSteelResult3 = SapModel.DesignResults.DesignForces.BraceDesignForces(str(u))
        try:
            MaxDesignSteelResult = max([abs(t) for t in DesignSteelResult3[4]])
        except:
            pass
        
        AxialForceList.append(MaxDesignSteelResult)
    MaxPu = max(AxialForceList)
    [PropName, SAuto, ret] = SapModel.FrameObj.GetSection(str(BrEcl.loc[i, 'ObjectName']))
    gyr = FrameSec_steelTube.loc[PropName, "R33"]
    BrKLR = 1*RealArcLength/gyr
    BrKLRLimit = 4.71*(200000/eFy_A500GrB)**0.5
    Fe = (math.pi)**2*200000/BrKLR**2
    if BrKLR <= BrKLRLimit:
        Fcr = (0.658**(eFy_A500GrB/Fe))*eFy_A500GrB
    if BrKLR > BrKLRLimit:
        Fcr = 0.877*Fe
    FiPnt = 0.9*eFy_A500GrB*FrameSec_steelTube.loc[PropName, "Area"]
    FiPnc = 0.9*Fcr*FrameSec_steelTube.loc[PropName, "Area"]
    
    BrEcl.loc[i, 'PutoFiPn'] = max(MaxPu/FiPnt, MaxPu/FiPnc)
    BrEcl.loc[i+1, 'PutoFiPn'] = max(MaxPu/FiPnt, MaxPu/FiPnc)

SapModel.SetModelIsLocked(False)
for i in range(len(BrEcl)) : 
    [PropName, SAuto, ret] = SapModel.FrameObj.GetSection(str(BrEcl.loc[i, 'ObjectName']))
    if BrEcl.loc[i, 'PutoFiPn'] < 0.2:
        Ecl_Pn = 1/((1/(2*0.9*1.14*Fcr*FrameSec_steelTube.loc[PropName, "Area"]))+(e/(0.9*(FrameSec_steelTube.loc[PropName, "Z33"])*eFy_A500GrB)))
        Ecl_03Pn = 1/((1/(2*0.9*0.3*1.14*Fcr*FrameSec_steelTube.loc[PropName, "Area"]))+(e/(0.9*(FrameSec_steelTube.loc[PropName, "Z33"])*eFy_A500GrB)))
        Ecl_tn = 1/((1/(2*0.9*eFy_A500GrB*FrameSec_steelTube.loc[PropName, "Area"]))+(e/(0.9*(FrameSec_steelTube.loc[PropName, "Z33"])*eFy_A500GrB)))
    
    elif BrEcl.loc[i, 'PutoFiPn'] >= 0.2:
        Ecl_Pn = 1/((1/(0.9*1.14*Fcr*FrameSec_steelTube.loc[PropName, "Area"]))+(8/9*e/(0.9*(FrameSec_steelTube.loc[PropName, "Z33"])*eFy_A500GrB)))
        Ecl_03Pn = 1/((1/(0.9*0.3*1.14*Fcr*FrameSec_steelTube.loc[PropName, "Area"]))+(8/9*e/(0.9*(FrameSec_steelTube.loc[PropName, "Z33"])*eFy_A500GrB)))
        Ecl_tn = 1/((1/(0.9*eFy_A500GrB*FrameSec_steelTube.loc[PropName, "Area"]))+(8/9*e/(0.9*(FrameSec_steelTube.loc[PropName, "Z33"])*eFy_A500GrB)))
    
    index = FrameList.index(str(BrEcl.loc[i, 'ObjectName']))
    if BrEcl.loc[i, 'BrLoc'] == 'End I-L' :
        BrEcl.loc[i, 'Ecl_Tn'] = Ecl_tn
        # SapModel.PointObj.SetLoadForce(Name, LoadPat, Value, Replace, CSys, ItemType)
        SapModel.PointObj.SetLoadForce(GetAllFrames[4][index], "Ecl (Tn)", [+Ecl_tn*(h/straight_leng), 0, +Ecl_tn*(b/2/straight_leng), 0, 0, 0], True, "Global", 0)
        
        # # SapModel.FrameObj.SetLoadPoint(Name, LoadPat, MyType, Dir, Dist, Val, CSys, RelDist, Replace, ItemType)
        # SapModel.FrameObj.SetLoadPoint(str(BrEcl.loc[i, 'ObjectName']), "Ecl (Tn)", 1, 4, 0, -Ecl_tn*(h/straight_leng), "Global", True, True, 0)
        # SapModel.FrameObj.SetLoadPoint(str(BrEcl.loc[i, 'ObjectName']), "Ecl (Tn)", 1, 6, 0, -Ecl_tn*(b/2/straight_leng), "Global", True, False, 0)
    
    if BrEcl.loc[i, 'BrLoc'] == 'End J-L' :
        BrEcl.loc[i, 'Ecl_Tn'] = Ecl_tn
        # SapModel.PointObj.SetLoadForce(Name, LoadPat, Value, Replace, CSys, ItemType)
        SapModel.PointObj.SetLoadForce(GetAllFrames[5][index], "Ecl (Tn)", [-Ecl_tn*(h/straight_leng), 0, -Ecl_tn*(b/2/straight_leng), 0, 0, 0], True, "Global", 0)
        
        # # SapModel.FrameObj.SetLoadPoint(Name, LoadPat, MyType, Dir, Dist, Val, CSys, RelDist, Replace, ItemType)
        # SapModel.FrameObj.SetLoadPoint(str(BrEcl.loc[i, 'ObjectName']), "Ecl (Tn)", 1, 4, 1, +Ecl_tn*(h/straight_leng), "Global", True, True, 0)
        # SapModel.FrameObj.SetLoadPoint(str(BrEcl.loc[i, 'ObjectName']), "Ecl (Tn)", 1, 6, 1, +Ecl_tn*(b/2/straight_leng), "Global", True, False, 0)
    
    if BrEcl.loc[i, 'BrLoc'] == 'End I-R' :
        BrEcl.loc[i, 'Ecl_Pn'] = Ecl_Pn
        BrEcl.loc[i, 'Ecl_03Pn'] = Ecl_03Pn
        # SapModel.PointObj.SetLoadForce(Name, LoadPat, Value, Replace, CSys, ItemType)
        SapModel.PointObj.SetLoadForce(GetAllFrames[4][index], "Ecl (Pn)", [+Ecl_Pn*(h/straight_leng), 0, -Ecl_Pn*(b/2/straight_leng), 0, 0, 0], True, "Global", 0)
        SapModel.PointObj.SetLoadForce(GetAllFrames[4][index], "Ecl (0.3Pn)", [+Ecl_03Pn*(h/straight_leng), 0, -Ecl_03Pn*(b/2/straight_leng), 0, 0, 0], True, "Global", 0)
        
        # # SapModel.FrameObj.SetLoadPoint(Name, LoadPat, MyType, Dir, Dist, Val, CSys, RelDist, Replace, ItemType)
        # SapModel.FrameObj.SetLoadPoint(str(BrEcl.loc[i, 'ObjectName']), "Ecl (Pn)", 1, 4, 0, -Ecl_Pn*(h/straight_leng), "Global", True, True, 0)
        # SapModel.FrameObj.SetLoadPoint(str(BrEcl.loc[i, 'ObjectName']), "Ecl (Pn)", 1, 6, 0, +Ecl_Pn*(b/2/straight_leng), "Global", True, False, 0)
        # SapModel.FrameObj.SetLoadPoint(str(BrEcl.loc[i, 'ObjectName']), "Ecl (0.3Pn)", 1, 4, 0, -Ecl_03Pn*(h/straight_leng), "Global", True, True, 0)
        # SapModel.FrameObj.SetLoadPoint(str(BrEcl.loc[i, 'ObjectName']), "Ecl (0.3Pn)", 1, 6, 0, +Ecl_03Pn*(b/2/straight_leng), "Global", True, False, 0)
    
    if BrEcl.loc[i, 'BrLoc'] == 'End J-R' :
        BrEcl.loc[i, 'Ecl_Pn'] = Ecl_Pn
        BrEcl.loc[i, 'Ecl_03Pn'] = Ecl_03Pn
        # SapModel.PointObj.SetLoadForce(Name, LoadPat, Value, Replace, CSys, ItemType)
        SapModel.PointObj.SetLoadForce(GetAllFrames[5][index], "Ecl (Pn)", [-Ecl_Pn*(h/straight_leng), 0, +Ecl_Pn*(b/2/straight_leng), 0, 0, 0], True, "Global", 0)
        SapModel.PointObj.SetLoadForce(GetAllFrames[5][index], "Ecl (0.3Pn)", [-Ecl_03Pn*(h/straight_leng), 0, +Ecl_03Pn*(b/2/straight_leng), 0, 0, 0], True, "Global", 0)
        
        # # SapModel.FrameObj.SetLoadPoint(Name, LoadPat, MyType, Dir, Dist, Val, CSys, RelDist, Replace, ItemType)
        # SapModel.FrameObj.SetLoadPoint(str(BrEcl.loc[i, 'ObjectName']), "Ecl (Pn)", 1, 4, 1, +Ecl_Pn*(h/straight_leng), "Global", True, True, 0)
        # SapModel.FrameObj.SetLoadPoint(str(BrEcl.loc[i, 'ObjectName']), "Ecl (Pn)", 1, 6, 1, -Ecl_Pn*(b/2/straight_leng), "Global", True, False, 0)
        # SapModel.FrameObj.SetLoadPoint(str(BrEcl.loc[i, 'ObjectName']), "Ecl (0.3Pn)", 1, 4, 1, +Ecl_03Pn*(h/straight_leng), "Global", True, True, 0)
        # SapModel.FrameObj.SetLoadPoint(str(BrEcl.loc[i, 'ObjectName']), "Ecl (0.3Pn)", 1, 6, 1, -Ecl_03Pn*(b/2/straight_leng), "Global", True, False, 0)

save_vars = "[BrEcl]"
save_vars = save_vars.strip('[]').split(',')
save_vars = [element.strip() for element in save_vars]
save_variables(f"{file_path}/Database.json", *save_vars)

#%% Assign Brace modifiers

Value = [0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 0.000001, 1.0, 1.0]
SapModel.FrameObj.SetModifiers("Braces (ALL)", Value, 1)

#%% Assign Story Restraints

# Restraint = [U1, U2, U3, R1, R2, R3]
Restraint = [True, False, False, False, False, False]
# SapModel.PointObj.SetRestraint(Name, Value, ItemType)
SapModel.PointObj.SetRestraint("Ecl Fixed Nodes", Restraint, 1)

#%% Set new Load Cases

# SapModel.LoadCases.StaticLinear.SetCase(Name)
SapModel.LoadCases.StaticLinear.SetCase("Ecl (Tn) + Ecl (Pn)")
SapModel.LoadCases.StaticLinear.SetCase("Ecl (Tn) + Ecl (0.3Pn)")

# SapModel.LoadCases.StaticLinear.SetLoads(Name, NumberLoads, LoadType, LoadName, SF)
SapModel.LoadCases.StaticLinear.SetLoads("Ecl (Tn) + Ecl (Pn)", 2, ["Load", "Load"], ["Ecl (Tn)", "Ecl (Pn)"], [1.0, 1.0])
SapModel.LoadCases.StaticLinear.SetLoads("Ecl (Tn) + Ecl (0.3Pn)", 2, ["Load", "Load"], ["Ecl (Tn)", "Ecl (0.3Pn)"], [1.0, 1.0])

#%% Reset Brace modifiers

Value = [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0]
SapModel.FrameObj.SetModifiers("Braces (ALL)", Value, 1)

#%% Reset Story Restraints

# Restraint = [U1, U2, U3, R1, R2, R3]
Restraint = [False, False, False, False, False, False]
# SapModel.PointObj.SetRestraint(Name, Value, ItemType)
SapModel.PointObj.SetRestraint("Ecl Fixed Nodes", Restraint, 1)

#%% Add fake dead loads

Type_Other = 8
# SapModel.LoadPatterns.Add(Name, MyType)
SapModel.LoadPatterns.Add("Dead (Beam)", Type_Other)
SapModel.LoadPatterns.Add("Dead (Column)", Type_Other)
SapModel.LoadPatterns.Add("Dead (Brace)", Type_Other)

# SapModel.FrameObj.SetLoadDistributed(Name, LoadPat, MyType, Dir, Dist1, Dist2, Val1, Val2, CSys, RelDist, Replace, ItemType)
SapModel.FrameObj.SetLoadDistributed("Beams (ALL)", "Dead (Beam)", 1, 10, 0, 1, 1, 1, "Global", True, True, 1)

# SapModel.FrameObj.SetLoadDistributed(Name, LoadPat, MyType, Dir, Dist1, Dist2, Val1, Val2, CSys, RelDist, Replace, ItemType)
SapModel.FrameObj.SetLoadDistributed("Columns (ALL)", "Dead (Column)", 1, 10, 0, 1, 1, 1, "Global", True, True, 1)

#%% Add UDStlPzPg Load Combination

SapModel.RespCombo.Add("UDStlPzPg", 0)
SapModel.RespCombo.SetCaseList("UDStlPzPg", 0, "Dead", 1.05)
SapModel.RespCombo.SetCaseList("UDStlPzPg", 0, "SuperDead", 1.05)
SapModel.RespCombo.SetCaseList("UDStlPzPg", 0, "Slab", 1.05)
SapModel.RespCombo.SetCaseList("UDStlPzPg", 0, "Partition", 1.05)
SapModel.RespCombo.SetCaseList("UDStlPzPg", 0, "Wall", 1.05)
SapModel.RespCombo.SetCaseList("UDStlPzPg", 0, "L1D", 1.05)
SapModel.RespCombo.SetCaseList("UDStlPzPg", 0, "Live", 0.25)
SapModel.RespCombo.SetCaseList("UDStlPzPg", 0, "L1L", 0.25)

#%% Set Insertion Points

SapModel.FrameObj.SetInsertionPoint("Beams (ALL)", 10, False, False, [0, 0, 0], [0, 0, 0], "Global", 1)

#%% Correct Mass Source

LoadPat = ["Dead", "SuperDead", "Slab", "Partition", "Wall", "L2D", "Live", "L2L"]
SF = [1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 0.25, 0.25]
# SapModel.PropMaterial.SetMassSource(MyOption, NumberLoads, LoadPat, sf)
SapModel.PropMaterial.SetMassSource(3, len(LoadPat), LoadPat, SF)

#%% Add End Offset joints

[NumberItems, ObjectType, ObjectName, ret] = SapModel.GroupDef.GetAssignments("Columns (ALL)")
ColumnsList = pd.DataFrame()
ColumnsList ['ObjectName'] = ObjectName
ColumnsList ['ObjectName'] = ColumnsList ['ObjectName'].astype(int)
ColumnEndI = []
ColumnEndJ = []
ColumnSection = []
ColumnDepth = []
for i in ColumnsList ['ObjectName']:
    ColumnEnds = SapModel.FrameObj.GetPoints(str(i))
    ColumnEndI.append(ColumnEnds[0])
    ColumnEndJ.append(ColumnEnds[1])
    ColumnSec = SapModel.FrameObj.GetSection(str(i))
    ColumnSection.append(ColumnSec[0])
    ColumnSecProp = SapModel.PropFrame.GetISection(ColumnSec[0])
    ColumnDepth.append(ColumnSecProp[2])
ColumnsList ['EndI'] = ColumnEndI
ColumnsList ['EndJ'] = ColumnEndJ
ColumnsList ['Section'] = ColumnSection
ColumnsList ['Total Depth'] = ColumnDepth
ColumnsList ['EndI'] = ColumnsList ['EndI'].astype(int)
ColumnsList ['EndJ'] = ColumnsList ['EndJ'].astype(int)

[NumberItems, ObjectType, ObjectName, ret] = SapModel.GroupDef.GetAssignments("Beams (ALL)")
BeamsList = pd.DataFrame()
BeamsList ['ObjectName'] = ObjectName
BeamsList ['ObjectName'] = BeamsList ['ObjectName'].astype(int)
ColumnEndI = []
ColumnEndJ = []
Beamsection = []
ColumnDepth = []
for i in BeamsList ['ObjectName']:
    ColumnEnds = SapModel.FrameObj.GetPoints(str(i))
    ColumnEndI.append(ColumnEnds[0])
    ColumnEndJ.append(ColumnEnds[1])
    Beamsec = SapModel.FrameObj.GetSection(str(i))
    Beamsection.append(Beamsec[0])
    BeamsecProp = SapModel.PropFrame.GetISection(Beamsec[0])
    ColumnDepth.append(BeamsecProp[2])
BeamsList ['EndI'] = ColumnEndI
BeamsList ['EndJ'] = ColumnEndJ
BeamsList ['Section'] = Beamsection
BeamsList ['Total Depth'] = ColumnDepth
BeamsList ['EndI'] = BeamsList ['EndI'].astype(int)
BeamsList ['EndJ'] = BeamsList ['EndJ'].astype(int)

EndLengthColI = []
EndLengthColJ = []
for i in range(len(ColumnsList)):
    indexI = BeamsList[BeamsList['EndI'] == int(ColumnsList.loc[i, 'EndI'])].index.tolist()
    indexJ = BeamsList[BeamsList['EndJ'] == int(ColumnsList.loc[i, 'EndI'])].index.tolist()
    if indexI:
        EndIDepth = BeamsList.loc[indexI, "Total Depth"]
        EndIDepth = float(EndIDepth.iloc[0])
    else:
        EndIDepth = 0
    if indexJ:
        EndJDepth = BeamsList.loc[indexJ, "Total Depth"]
        EndJDepth = float(EndJDepth.iloc[0])
    else:
        EndJDepth = 0
    EndLengthColI.append(max(EndIDepth, EndJDepth)/2)
    
    indexI = BeamsList[BeamsList['EndI'] == int(ColumnsList.loc[i, 'EndJ'])].index.tolist()
    indexJ = BeamsList[BeamsList['EndJ'] == int(ColumnsList.loc[i, 'EndJ'])].index.tolist()
    if indexI:
        EndIDepth = BeamsList.loc[indexI, "Total Depth"]
        EndIDepth = float(EndIDepth.iloc[0])
    else:
        EndIDepth = 0
    if indexJ:
        EndJDepth = BeamsList.loc[indexJ, "Total Depth"]
        EndJDepth = float(EndJDepth.iloc[0])
    else:
        EndJDepth = 0
    EndLengthColJ.append(max(EndIDepth, EndJDepth)/2)
ColumnsList ['EndLenI'] = EndLengthColI
ColumnsList ['EndLenJ'] = EndLengthColJ

EndLengthBeamI = []
EndLengthBeamJ = []
for i in range(len(BeamsList)):
    indexI = ColumnsList[ColumnsList['EndI'] == int(BeamsList.loc[i, 'EndI'])].index.tolist()
    indexJ = ColumnsList[ColumnsList['EndJ'] == int(BeamsList.loc[i, 'EndI'])].index.tolist()
    if indexI:
        EndIDepth = ColumnsList.loc[indexI, "Total Depth"]
        EndIDepth = float(EndIDepth.iloc[0])
    else:
        EndIDepth = 0
    if indexJ:
        EndJDepth = ColumnsList.loc[indexJ, "Total Depth"]
        EndJDepth = float(EndJDepth.iloc[0])
    else:
        EndJDepth = 0
    EndLengthBeamI.append(max(EndIDepth, EndJDepth)/2)
    
    indexI = ColumnsList[ColumnsList['EndI'] == int(BeamsList.loc[i, 'EndJ'])].index.tolist()
    indexJ = ColumnsList[ColumnsList['EndJ'] == int(BeamsList.loc[i, 'EndJ'])].index.tolist()
    if indexI:
        EndIDepth = ColumnsList.loc[indexI, "Total Depth"]
        EndIDepth = float(EndIDepth.iloc[0])
    else:
        EndIDepth = 0
    if indexJ:
        EndJDepth = ColumnsList.loc[indexJ, "Total Depth"]
        EndJDepth = float(EndJDepth.iloc[0])
    else:
        EndJDepth = 0
    EndLengthBeamJ.append(max(EndIDepth, EndJDepth)/2)
BeamsList ['EndLenI'] = EndLengthBeamI
BeamsList ['EndLenJ'] = EndLengthBeamJ

for i in range(len(BeamsList)):
    EndI = BeamsList.loc[i, 'EndI']
    EndICoord = SapModel.PointObj.GetCoordCartesian(str(EndI))
    SapModel.PointObj.AddCartesian(EndICoord[0]+BeamsList.loc[i, 'EndLenI'], EndICoord[1], EndICoord[2])
    
    EndJ = BeamsList.loc[i, 'EndJ']
    EndJCoord = SapModel.PointObj.GetCoordCartesian(str(EndJ))
    SapModel.PointObj.AddCartesian(EndJCoord[0]-BeamsList.loc[i, 'EndLenJ'], EndJCoord[1], EndJCoord[2])
    
for i in range(len(ColumnsList)):
    EndI = ColumnsList.loc[i, 'EndI']
    EndICoord = SapModel.PointObj.GetCoordCartesian(str(EndI))
    SapModel.PointObj.AddCartesian(EndICoord[0], EndICoord[1], EndICoord[2]+ColumnsList.loc[i, 'EndLenI'])
    
    EndJ = ColumnsList.loc[i, 'EndJ']
    EndJCoord = SapModel.PointObj.GetCoordCartesian(str(EndJ))
    SapModel.PointObj.AddCartesian(EndJCoord[0], EndJCoord[1], EndJCoord[2]-ColumnsList.loc[i, 'EndLenJ'])

#%% Rigid Modifiers 

StiffPropModif = [10, 1, 1, 1, 10, 10, 1, 1]
SapModel.FrameObj.SetModifiers("Columns (Stiff)", StiffPropModif, 1)
SapModel.FrameObj.SetModifiers("Beams (Stiff)", StiffPropModif, 1)
SapModel.FrameObj.SetModifiers("Braces (Stiff)", StiffPropModif, 1)

PropModif = [1, 1, 1, 1, 1, 1, 1, 1]
SapModel.FrameObj.SetModifiers("Columns (ALL)", PropModif, 1)
SapModel.FrameObj.SetModifiers("Beams (ALL)", PropModif, 1)
SapModel.FrameObj.SetModifiers("Braces (ALL)", PropModif, 1)

#%% Export Tables to excel

N_mm_C = 9
SapModel.SetPresentUnits(N_mm_C)

TableKey = ["Assembled Joint Masses",
            "Beam Object Connectivity", "Brace Object Connectivity", "Column Object Connectivity",
            "Element Forces - Columns",
            "Frame Assignments - Releases and Partial Fixity", "Frame Assignments - Section Properties",
            "Frame Loads Assignments - Distributed", "Frame Section Property Definitions - Summary",
            "Frame Section Property Definitions - Steel I/Wide Flange", "Frame Section Property Definitions - Steel Tube",
            "Group Assignments", "Joint Assignments - Restraints", "Joint Loads Assignments - Force",
            "Material Properties - Steel Data", "Material Properties - User Stress-Strain Curves", "Modal Participating Mass Ratios", "Point Object Connectivity"]
SapModel.DatabaseTables.ShowTablesInExcel(TableKey, 0)

