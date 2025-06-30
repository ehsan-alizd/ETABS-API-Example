#%% Script data

# Ehsan Alizadeh
# LinkedIn: https://www.linkedin.com/in/ehsanalizd

#%% Input

file_path = r"D:\Thesis"

#%% Import Libraries

import os
import sys
import comtypes.client
import inspect
import json
import numpy as np
import pandas as pd

#%% Define Functions

file_path = file_path.replace("\\", "/")

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

#initialize model
SapModel.InitializeNewModel()

#create new blank model
ret = SapModel.File.NewBlank()

#%% Define materials

## Switch units
kN_mm_C = 5
kN_m_C = 6
kgf_mm_C = 7
kgf_m_C = 8
N_mm_C = 9
N_m_C = 10
SapModel.SetPresentUnits(N_mm_C)

# ## Define material property
# eMatType_Steel = 1

# # SapModel.PropMaterial.SetMaterial(Name, Mat_Type)
# SapModel.PropMaterial.SetMaterial("A992", eMatType_Steel)
# SapModel.PropMaterial.SetMaterial("A500GrB", eMatType_Steel)

# # SapModel.PropMaterial.SetMPIsotropic(Name, The modulus of elasticity, Poisson’s ratio, The thermal coefficient)
# SapModel.PropMaterial.SetMPIsotropic("A992", 200000, 0.3, 0.0000117)
# SapModel.PropMaterial.SetMPIsotropic("A500GrB", 200000, 0.3, 0.0000117)

Fy_A992 = 345
Fu_A992 = 450
eFy_A992 = 379.5
eFu_A992 = 495
Fy_A500GrB = 315
Fu_A500GrB = 400
eFy_A500GrB = 441
eFu_A500GrB = 520

# # SapModel.PropMaterial.SetOSteel_1(Name, Fy, Fu, eFy, eFu, SSType, SSHysType, StrainAtHardening, StrainAtMaxStress, StrainAtRupture, FinalSlope)
# SapModel.PropMaterial.SetOSteel_1("A992", Fy_A992, Fu_A992, eFy_A992, eFu_A992, 1, 2, 0.02, 0.1, 0.2, -0.1)
# SapModel.PropMaterial.SetOSteel_1("A500GrB", Fy_A500GrB, Fu_A500GrB, eFy_A500GrB, eFu_A500GrB, 1, 2, 0.02, 0.1, 0.2, -0.1)

save_vars = "[Fy_A992, Fu_A992, eFy_A992, eFu_A992, Fy_A500GrB, Fu_A500GrB, eFy_A500GrB, eFu_A500GrB]"
save_vars = save_vars.strip('[]').split(',')
save_vars = [element.strip() for element in save_vars]
save_variables(f"{file_path}/Database.json", *save_vars)

#%% Define Load Patterns

Type_Dead = 1
Type_SuperSead = 2
Type_Live = 3
Type_Seismic = 5
Type_Other = 8
Type_Notional = 12

# SapModel.LoadPatterns.Add(Name, MyType)
SapModel.LoadPatterns.Add("SuperDead", Type_SuperSead)
SapModel.LoadPatterns.Add("Slab", Type_SuperSead)
SapModel.LoadPatterns.Add("Partition", Type_SuperSead)
SapModel.LoadPatterns.Add("Wall", Type_SuperSead)
SapModel.LoadPatterns.Add("Live", Type_Live)
SapModel.LoadPatterns.Add("L1D", Type_SuperSead)
SapModel.LoadPatterns.Add("L1L", Type_Live)
SapModel.LoadPatterns.Add("L2D", Type_Other)
SapModel.LoadPatterns.Add("L2L", Type_Other)
SapModel.LoadPatterns.Add("EX", Type_Seismic)
SapModel.LoadPatterns.Add("Ecl (Tn)", Type_Other)
SapModel.LoadPatterns.Add("Ecl (Pn)", Type_Other)
SapModel.LoadPatterns.Add("Ecl (0.3Pn)", Type_Other)

SapModel.LoadPatterns.Add("NDeadX", Type_Notional)
SapModel.LoadPatterns.Add("NSuperDeadX", Type_Notional)
SapModel.LoadPatterns.Add("NSlabX", Type_Notional)
SapModel.LoadPatterns.Add("NPartitionX", Type_Notional)
SapModel.LoadPatterns.Add("NWallX", Type_Notional)
SapModel.LoadPatterns.Add("NLiveX", Type_Notional)

#%% Define Mass Source

LoadPat = ["Dead", "SuperDead", "Slab", "Partition", "Wall", "L2D", "Live", "L2L"]
SF = [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 0.2, 0.2]
# SapModel.PropMaterial.SetMassSource(MyOption, NumberLoads, LoadPat, sf)
SapModel.PropMaterial.SetMassSource(3, len(LoadPat), LoadPat, SF)

#%% Define Groups

SapModel.GroupDef.SetGroup("Columns (ALL)")
SapModel.GroupDef.SetGroup("Columns (Braced bay)")
SapModel.GroupDef.SetGroup("Columns (Gravity)")
SapModel.GroupDef.SetGroup("Columns (Stiff)")

SapModel.GroupDef.SetGroup("Beams (ALL)")
SapModel.GroupDef.SetGroup("Beams (Braced bay)")
SapModel.GroupDef.SetGroup("Beams (Gravity)")
SapModel.GroupDef.SetGroup("Beams (Stiff)")

SapModel.GroupDef.SetGroup("Braces (ALL)")
SapModel.GroupDef.SetGroup("Braces (Stiff)")
SapModel.GroupDef.SetGroup("Braces (End I)")
SapModel.GroupDef.SetGroup("Braces (End I-R)")
SapModel.GroupDef.SetGroup("Braces (End I-L)")
SapModel.GroupDef.SetGroup("Braces (End J)")
SapModel.GroupDef.SetGroup("Braces (End J-R)")
SapModel.GroupDef.SetGroup("Braces (End J-L)")

SapModel.GroupDef.SetGroup("Ecl Fixed Nodes")
SapModel.GroupDef.SetGroup("Leaning Column")
SapModel.GroupDef.SetGroup("Leaning Beam")
SapModel.GroupDef.SetGroup("Leaning Joint")

SapModel.GroupDef.SetGroup("Panel Zone")
