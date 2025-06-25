import tkinter as tk
import numpy as np
import math
import pandas as pd
from tkinter import filedialog, messagebox, ttk
from tkinter import *
from tkinter.ttk import *
import re
import os

# Get the directory of the current Python script
script_directory = os.path.dirname(os.path.abspath(__file__))

# Set the current working directory to the directory of the script
os.chdir(script_directory)

def createDF():
    global df
    df = pd.DataFrame(columns=['Code', 'Description', 'Quantity', 'Price', 'Discount', 'Discount Price', 'Total'])

def Refresh():
    global df
    Emptentry = pd.DataFrame({"Code": [" "],
                            "Description": [" "],
                            "Quantity": [" "],
                            "Price": [" "],
                            "Discount": [" "],
                            "Discount Price": [" "],
                            "Total": [" "]})
    df = pd.concat([df, Emptentry])
    
    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text = column)
        
    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tv1.insert("", "end", values=row)
    
    tv1.column("Code", width = 80)
    tv1.column("Description", width = 350)
    tv1.column("Quantity", width = 50, anchor=tk.CENTER)
    tv1.column("Price", width = 65, anchor=tk.CENTER)
    tv1.column("Discount", width = 40, anchor=tk.CENTER)
    tv1.column("Discount Price", width = 65, anchor=tk.CENTER)
    tv1.column("Total", width = 80, anchor=tk.CENTER)
    return None

def File_dialog():
    filename = filedialog.askopenfilename(initialdir="/", 
                                        title="Select A File", 
                                        filetypes=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file["text"] = filename
    return None
    
def MemberList():
    global df
    global Costdf
    global RafterList
    global Rafterdf
    global Purlindf
    global PurlinList
    global Piledf
    global PileList
    global SBdf
    global SBList

    Costdf = pd.read_excel("Member Rates.xlsx")

    Rafterdf = Costdf.loc[:, ['Rafter Code', 'Rafter Description', 'Rafter Weight [kg/m]']]
    Rafterdf = Rafterdf.dropna()
    RafterList = Rafterdf.iloc[:,0].tolist()

    Purlindf = Costdf.loc[:, ['Purlin Code', 'Purlin Description', 'Purlin Weight [kg/m]']]
    Purlindf = Purlindf.dropna()
    
    Piledf = Costdf.loc[:, ['Pile Code', 'Pile Description', 'Pile Weight [kg/m]']]
    Piledf = Piledf.dropna()
    PileList = Piledf.iloc[:, 1].tolist()
    
    # Create shortened display versions and a mapping to full descriptions
    global PileShortList, ShortToFullMap
    PileShortList = []
    ShortToFullMap = {}

    for desc in PileList:
        # Try to extract dimension (e.g., after "DIM:")
        match = re.search(r'DIM:([0-9x]+mm)', desc)
        if match:
            short = match.group(1)
        else:
            short = desc  # Fallback

        PileShortList.append(short)
        ShortToFullMap[short] = desc

    SBdf = Costdf.loc[:, ['Support Bar Code', 'Support Bar Description', 'Support Bar Weight [kg/m]']]
    SBdf = SBdf.dropna()
    SBList = SBdf.iloc[:,1].tolist()

def Load_excel_data():
    File_dialog()
    MemberList()

    file_path = label_file["text"]
    global pricedf
    try:
        excel_filename = r"{}".format(file_path)
        pricedf = pd.read_excel(excel_filename)
        pricedf = pricedf.iloc[2:, 0:3]
    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid.")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None
    label_file.config(text = "Load successful")
    
    print(pricedf.head())
    
def getCustomerList():
    global CIDList
    CIDList = Customerdf.iloc[:,0]
    global CNameList
    CNameList = Customerdf.iloc[:,1]

    global CList
    CList = []

    for i in range(0, len(CIDList)):
        CID = str(CIDList[i])
        CName = str(CNameList[i])
        
        CString = CID + " - " + CName
        
        CList.append(CString)
    
def Load_Customer_excel_data():
    File_dialog()

    file_path = label_file["text"]
    global Customerdf
    try:
        excel_filename = r"{}".format(file_path)
        Customerdf = pd.read_excel(excel_filename)
        Customerdf = Customerdf.iloc[:, [0, 12]]
        getCustomerList()
        updateListBox(CList)
    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid.")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None
    
    label_file.config(text = "Load successful")
    print(Customerdf.head())
    
def Save_Excel():
    global df
    global K8df
    global quote_weight_df
    
    ConvertToK8()
    CreateWeightDF()
    
    CombinedDF = CombineDataFrames(df, K8df)
    CombinedDF = CombineDataFrames(CombinedDF, quote_weight_df)
    
    file = filedialog.asksaveasfilename(defaultextension = ".xlsx")
    CombinedDF.to_excel(str(file))
    label_file.config(text = "File saved")

def clear_data():
    tv1.delete(*tv1.get_children())
    pass
    
def getMount():
    global Mount
    global MountS
    
    MountS = MountVar.get()
    
    if (MountS == 'Concrete Ground Mount'):
        Mount = 0
    elif (MountS == 'Ground Screw'):
        Mount = 1
    elif (MountS == 'Simple Piles'):
        Mount = 2
    elif (MountS == 'Y-structure Piles'):
        Mount = 3

def GetPileOp():
    global Weightpm
    global PileSOption
    global Piledf
    global PileList  
    global WeightpmList
    global PileDescr
    global PileCode

    selected_short = PileSVar.get()  # e.g., '175x75x20x3mm'
    PileSOption = ShortToFullMap[selected_short]

    WeightpmList = Piledf.iloc[:, 2]
    index = PileList.index(PileSOption)

    Weightpm = WeightpmList[index]
    PileDescr = Piledf.iloc[index, 1]
    PileCode = Piledf.iloc[index, 0]

def round_up(number):
    """
    Rounds a number up to the nearest integer.

    Args:
        number (float): The number to round up.

    Returns:
        int: The rounded up integer.
    """
    return math.ceil(number)

def extract_percentage_value(percentage_str):
    """
    Extracts the numeric value from a percentage string and returns it as an integer.

    Args:
        percentage_str (str): A string containing a percentage (e.g., "10%").

    Returns:
        int: The numeric value without the percentage sign.
    """
    return int(percentage_str.strip().replace('%', ''))

def getSmalls():
    global SmallSmalls
    global ConSmalls
    global SuppSmalls
    
    SmallSmalls = 1 + extract_percentage_value(SSmallsVar.get())/100
    ConSmalls = 1 + extract_percentage_value(ConSmallsVar.get())/100
    SuppSmalls = 1 + extract_percentage_value(SuppSmallsVar.get())/100

def getInputs():

    createDF()
    
    #get Numerical Inputs
    global sysnum
    global pHor
    global pWidth
    global pLength
    global GroundClearance
    global MaxSuppSpace
    global pNum
    global angle

    sysnum = int(TableNumberE.get())
    pHor = int(HorPanelE.get())
    pWidth = float(PanelWidthE.get())
    pLength = float(PanelLengthE.get())
    angle = np.deg2rad(float(AngleE.get()))
    GroundClearance = float(GroundClearanceE.get())
    MaxSuppSpace = float(MaxSuppSpaceE.get())

    #Vert Panels
    global pVert
    pVert = int(VertVar.get())
    selection = str(pVert)
    pNum = pVert*pHor

    #get ROH
    global ROH
    ROHS = var.get()
    
    if(ROHS == '600mm'):
        ROH = 600
    elif(ROHS == '800mm'):
        ROH = 800
    selection = str(ROH) 

    #get Discount 
    global discount
    discount = 1 - float(DiscountE.get())/100

    #Shared rails
    global ShR
    ShRS = ShRailVar.get()
    if(ShRS == "Yes"):
        ShR = 0
    else:
        ShR = 1
    selection = str(ShR)

    #Mount Selection
    global Mount
    global MountS
    
    MountS = MountVar.get()
    
    if (MountS == 'Concrete Ground Mount'):
        Mount = 0
    elif (MountS == 'Ground Screw'):
        Mount = 1
    elif (MountS == 'Simple Piles'):
        Mount = 2
    elif (MountS == 'Y-structure Piles'):
        Mount = 3
        
    #get Embedment Depth
    global EmbedmentD
    EmbedmentD = float(EmbedmentDE.get())
        
    #get Mark Up
    global MarkUp
    MU = MarkUpE.get()

    if (MU is None or MU == 0):
        MarkUp = 1
    else:
        MarkUp = float(MU)/100 + 1
        
    #get Pile Option and weight per meter for that Pile
    GetPileOp()
    
    #get Rand per meter for steel pile
    global Rate
    Rate = float(RateE.get())
    
    #get Smalls percentages and multiipliers
    getSmalls()
    
def getRaftChoice():
    global RafterLChosen
    global Rafterdf
    global RafterDescr
    global RafterCode
    
    RafterLChosenStr = str(RaftVar.get())
    
    RafterList = Rafterdf['Rafter Description'].tolist()
    
    index = 0
    
    for i in range(2,len(RafterList)):
            if (RafterList[i] == RafterLChosenStr):
                index = i
    
    RafterDescr = RafterList[index]
    
    RafterCode = Rafterdf.iloc[index, 0]
    
    Length = int(RafterDescr[11:15])
    
    RafterLChosen = Length

def Pileprice(length):
    global Weightpm
    global Rate
    global MarkUp
    
    PilePrice = Weightpm * length/1000 * Rate * MarkUp
    
    return round(PilePrice, 2)

def RailCalc():
    global pVert
    global ShR
    global ShRE
    global RailMult
    global WRailMult
    global sysnum
    global pHor
    global pWidth
    global pLength
    global MaxSuppSpace
    global pNum
    global df

    getInputs()
    getMount()
    #getMount()
    
    pNum = pVert*pHor

    if (pVert == 1):
        RailMult = 2
        WRailMult = 0

        ShRE = "No"

    elif(pVert == 2):
        if (ShR == 0): #There are shared rails
            RailMult = 2
            WRailMult = 1

            ShRE = "Yes"

        elif (ShR == 1): #No shared rails
            RailMult = 4
            WRailMult = 0

            ShRE = "No"

    elif(pVert == 3):
        if(ShR == 0): #There are shared rails
            RailMult = 4
            WRailMult = 1

            ShRE = "Yes"

        elif(ShR == 1): #No shared rails
            RailMult = 6
            WRailMult = 0

            ShRE = "No"
    
    elif (pVert == 4):
        if (ShR == 0): #There are shared rails
            RailMult = 4
            WRailMult = 2

            ShRE = "No"

        elif (ShR == 1): #No shared rails
            RailMult = 8
            WRailMult = 0

            ShRE = "Yes"

    elif (pVert == 6):
        if (ShR == 0): #There are shared rails
            RailMult = 8
            WRailMult = 2

            ShRE = "Yes"

        elif(ShR == 1): #No shared rails
            RailMult = 12
            WRailMult = 0

            ShRE = "No" 

def ClampCalc():

    RailCalc()

    global sysnum
    global pVert
    global pHor
    global ShR
    global ShRE
    global RailMult
    global df
    global pNum
    global CalcRafterL

    pNum = pVert*pHor
    
    #End Clamp Calculations + Entry

    ECMult = (pHor//20 + 1)
    
    EndClamps = (2*RailMult)*sysnum*ECMult * SmallSmalls
    AddEntry("LM-EC35-RNW", round_up(EndClamps), 0)

    #InterClamp Calculations + Entry
    InterClamps = ((RailMult*(pHor - 1)) + WRailMult*2*pHor)*sysnum * SmallSmalls
    AddEntry("LM-IC35-GP1-RNW", round_up(InterClamps), 0)

def calculate_supports(structure_length, max_overhang_pct, min_overhang_pct, max_support_spacing):
    min_ratio = min_overhang_pct / 100
    max_ratio = max_overhang_pct / 100

    supports = 2  # Start with the minimum number of supports

    while True:
        for ratio in np.linspace(min_ratio, max_ratio, num=100):
            spacing = structure_length / (supports - 1 + 2 * ratio)
            overhang = ratio * spacing

            if spacing <= max_support_spacing:
                return round(supports), round(spacing), round(overhang)

        supports += 1

def getPurlins():

    ClampCalc()

    global sysnum
    global pHor
    global pWidth
    global pLength
    global RafterLChosen
    global SupportLegs
    global pVert
    global SupportSpacingL
    global df
    global ShR
    global Mount
    global RailMult
    global WRailMult
    global ShRE
    global MountS
    global PurlinLMin

    CalcPurlinL = pHor*pWidth + 20*(pHor - 1) + 100
    
    CalcPurlLabel.config(text = "Required Purlin Length (mm): " +str(CalcPurlinL))
    
    maxnum = 39
    curnum = 0

    PurlinL = CalcPurlinL + 100
    global PurlinLMin
    #PurlinLMin = PurlinL
    #PurlinLCalc = PurlinL

    L5250num = CalcPurlinL//5250 - 1
    Purlin5250 = L5250num*5250
    ShortCalcPurlin = CalcPurlinL - Purlin5250
    PurlinLMint = ShortCalcPurlin + 5000
    
    
    L1 = 5250
    L2 = 4400
    L3 = 3300

    L1num = 0
    L2num = 0
    L3num = 0
    
    WL1num = 0
    WL2num = 0
    WL3num = 0
    
    L1numt = 0
    WL1numt = 0
    

    for i in range(0,maxnum):
        for j in range(0,maxnum):
            for k in range(0,maxnum):
                
                PurlinLCalc = i*L1 + j*L2 + k*L3
                
                if (PurlinLCalc<PurlinLMint and PurlinLCalc>ShortCalcPurlin):
                    WL1numt = (i)*WRailMult*sysnum
                    WL2num = j*WRailMult*sysnum
                    WL3num = k*WRailMult*sysnum
                    
                    
                    L1numt = (i)*RailMult*sysnum 
                    L2num = j*RailMult*sysnum 
                    L3num = k*RailMult*sysnum 
                    PurlinLMint = PurlinLCalc
                    
    PurlinLMin = Purlin5250 + PurlinLMint
    
    WL1num = L5250num*WRailMult*sysnum + WL1numt
    L1num = L5250num*RailMult*sysnum + L1numt
    
    if(L1num>0):
        AddEntry("LM-R112-5250", round_up(L1num * SuppSmalls), 0)
    
    if(L2num>0):
        AddEntry("LM-R112-4400", round_up(L2num * SuppSmalls), 0)
    
    if(L3num>0):
        AddEntry("LM-R112-3300", round_up(L3num * SuppSmalls), 0)
    
    PurlinSplicersNum = L1num + L2num + L3num - (RailMult * sysnum)
    AddEntry("LM-RS-I-R112-300", round_up(PurlinSplicersNum * SuppSmalls), 0)
    
    WPurlinSplicersNum = 0
    
    if (WRailMult > 0): #There are wide purlins
        
        if(WL1num>0):
            AddEntry("LM-R112-W-5250", round_up(WL1num * SuppSmalls), 0)
        if(WL2num>0):
            AddEntry("LM-R112-W-4400", round_up(WL2num * SuppSmalls), 0)
        if(WL3num>0):
            AddEntry("LM-R112-W-3300", round_up(WL3num * SuppSmalls), 0)
        
        WPurlinSplicersNum = WL1num + WL2num + WL3num - sysnum
        AddEntry("LM-RS-I-R112-W-300", round_up(WPurlinSplicersNum * SuppSmalls), 0)
        
    StitchingScrews = 8*(PurlinSplicersNum + WPurlinSplicersNum)
    AddEntry("FS-S-22X6-C4", round_up(StitchingScrews * SmallSmalls), 0)
   
    PurlinSuppString = "Supplied Purlin Length: " + str(PurlinLMin) + "mm"
    PurlinLabel.config(text = PurlinSuppString)
     
    StructLength = CalcPurlinL + 250
    MaxOHang = 30
    MinOHang = 20

    #Previous attempt at support spacing calculation
    global SupportLegsC
    global OHang
    global SupportSpacingL
    SupportLegsC, SupportSpacingL, OHang = calculate_supports(StructLength, MaxOHang, MinOHang, MaxSuppSpace)

    global SupportLegs
    SupportLegs = SupportLegsC * sysnum

    #Adding Purlin to Rafter Connectors
    global PRC
    totalrails = WRailMult + RailMult
    PRC = ((totalrails * SupportLegsC)*2)*sysnum
    AddEntry("LM-PRC", round_up(PRC * SmallSmalls), 0)
    
    #Display the required Rafter Length
    if(pVert == 4 or pVert == 6):
        CalcRafterL = pVert/2 * pLength + ((pVert - 1)*20) - ROH/2
    else:
        CalcRafterL = pVert * pLength + ((pVert - 1)*20) - ROH
    
    selection = "Calculated Rafter length: "+str(CalcRafterL)
    CalcRaftLabel.config(text = selection)
    
    SupportString = "Support spacing = " + str(SupportSpacingL) + "mm"
    SupportSLabel.config(text = SupportString)

    PurlinSuppString = "Supplied Purlin Length: " + str(PurlinLMin) + "mm"
    PurlinLabel.config(text = PurlinSuppString)
    
    SupportLegsStr = "Number of Support Legs per structure: " + str(SupportLegsC)
    SupportLegsLabel.config(text = SupportLegsStr)
    
    OHangString = "Purlin Overhang: "+str(OHang)+"mm"
    OHangLabel.config(text = OHangString)
    
    selection = "Calculated Rafter length: "+str(CalcRafterL)
    CalcRaftLabel.config(text = selection)

def MountSupp():

    global sysnum
    global Mount
    global SupportLegs
    global pVert
    global RafterLChosen
    global RailMult
    global df
    global Weightpm
    global MarkUp
    global SupportSpacingL
    global EmbedmentD
    global RafterLChosen
    global Rafterdf
    global RafterDescr
    global RafterCode
    global FrontSupport, BackSupport
    global BraceL, Bracewpm
    
    if(pVert == 4 or pVert == 6):
        RafterQuantity = SupportLegs * 2
    elif(pVert == 1):
        RafterQuantity = SupportLegs//2 + 1
    else:
        RafterQuantity = SupportLegs
    
    RafterQuantity = round_up(RafterQuantity * SuppSmalls)
    
    rdescription, rprice, rdiscprice, rtotal = getprice(RafterCode, RafterQuantity, 0)
    discountp = float(DiscountE.get())
    RaftEntry = pd.DataFrame({"Code": [RafterCode], 
                            "Description": [RafterDescr],
                            "Quantity": [RafterQuantity], 
                            "Price": [rprice],
                            "Discount": [str(discountp) + "%"],
                            "Discount Price": [rdiscprice],
                            "Total": [rtotal]})
    df = pd.concat([df, RaftEntry])    

    if (Mount == 0 or Mount == 1): #Concrete Ground Mount and Ground Screw
        
        FootPieceS = 0 
        FootPieceD = 0
        FootPieceT = 0

        if(pVert == 1 or (pVert == 2 and RafterLChosen < 4200)):

            if(pVert == 1):
                RafterLChosen = RafterLChosen/2

            SBConS = SupportLegs
            AddEntry("LM-TTC-90", round_up(SBConS * ConSmalls), 0)
        
            SBConD = SupportLegs
            AddEntry("LM-TTC-120", round_up(SBConD * ConSmalls), 0)
            
            FootPieceS = SupportLegs
            AddEntry("LM-FP-90", round_up(FootPieceS * ConSmalls), 0)
            
            FootPieceD = SupportLegs
            AddEntry("LM-FP-120", round_up(FootPieceD * ConSmalls), 0)
            
            
            if(RafterLChosen == 4000 and angle == np.deg2rad(20) and GroundClearance == 500):
                
                FrontSupport = 530
                CrossSupport = 2795
                BackSupport = 1560
                
                AddEntry("LM-SB-" + str(FrontSupport), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(CrossSupport), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(BackSupport), round_up(SupportLegs * SuppSmalls), 0)
                
            elif(RafterLChosen == 3800 and angle == np.deg2rad(20) and GroundClearance == 500):
                
                FrontSupport = 530
                CrossSupport = 2615
                BackSupport = 1500
                
                AddEntry("LM-SB-" + str(FrontSupport), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(CrossSupport), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(BackSupport), round_up(SupportLegs * SuppSmalls), 0)
            
            elif(RafterLChosen == 4200 and angle == np.deg2rad(25) and GroundClearance == 500):
                
                FrontSupport = 590
                CrossSupport = 2890
                BackSupport = 1960
                
                AddEntry("LM-SB-" + str(FrontSupport), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(CrossSupport), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(BackSupport), round_up(SupportLegs * SuppSmalls), 0)
                
            elif(RafterLChosen == 4000 and angle == np.deg2rad(25) and GroundClearance == 500):
                
                FrontSupport = 590
                CrossSupport = 2710
                BackSupport = 1875
                
                AddEntry("LM-SB-" + str(FrontSupport), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(CrossSupport), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(BackSupport), round_up(SupportLegs * SuppSmalls), 0)
                
            elif(RafterLChosen == 3800 and angle == np.deg2rad(25) and GroundClearance == 500):
                
                FrontSupport = 615
                CrossSupport = 2540
                BackSupport = 1815
                
                AddEntry("LM-SB-" + str(FrontSupport), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(CrossSupport), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(BackSupport), round_up(SupportLegs * SuppSmalls), 0)
                
            else:

                FOHang =  ((pVert*pLength+20*(pVert-1)) - RafterLChosen)/2
                FrontSupport = float(GroundClearanceE.get()) + FOHang*np.sin(angle)
                AddEntry("LM-SB-L", round_up(SupportLegs * SuppSmalls), round(FrontSupport))
                
                Hypotenuse =  RafterLChosen - 1000
                DistBetwFandB = Hypotenuse * np.cos(angle)
                ChangeElev = ((Hypotenuse)**2 - (DistBetwFandB)**2)**0.5
                
                CrossSupport = ((FrontSupport)**2 + (DistBetwFandB)**2)**0.5
                AddEntry("LM-SB-L", round_up(SupportLegs * SuppSmalls), round(CrossSupport))
                
                BackSupport = FrontSupport + ChangeElev
                AddEntry("LM-SB-L", round_up(SupportLegs * SuppSmalls), round(BackSupport))
                
                
                
        elif(pVert == 3 or (pVert == 2 and RafterLChosen >= 4200)):
            
            SBConS = SupportLegs*2
            AddEntry("LM-TTC-90", round_up(SBConS * ConSmalls), 0)
            
            SBConD = SupportLegs
            AddEntry("LM-TTC-120", round_up(SBConD * ConSmalls), 0)
            
            FootPieceD = SupportLegs*2
            AddEntry("LM-FP-120", round_up(FootPieceD * ConSmalls), 0)
            
            if(RafterLChosen == 6500 and angle == np.deg2rad(15) and GroundClearance == 500):
                
                FrontSupport = 450
                CrossSupport1 = 2820
                CrossSupport2 = 2820
                BackSupport = 1870
                
                AddEntry("LM-SB-" + str(FrontSupport), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(CrossSupport1), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(CrossSupport2), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(BackSupport), round_up(SupportLegs * SuppSmalls), 0)
                
            elif(RafterLChosen == 6200 and angle == np.deg2rad(15) and GroundClearance == 500):
                
                FrontSupport = 450
                CrossSupport1 = 2670
                CrossSupport2 = 2670
                BackSupport = 1800
                
                AddEntry("LM-SB-" + str(FrontSupport), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(CrossSupport1), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(CrossSupport2), round_up(SupportLegs * SuppSmalls), 0)
                AddEntry("LM-SB-" + str(BackSupport), round_up(SupportLegs * SuppSmalls), 0)
                
            else:
                
                FOHang =  ((pVert*pLength+20*(pVert-1)) - RafterLChosen)/2
                FrontSupport = float(GroundClearanceE.get()) + FOHang*np.sin(angle)
                AddEntry("LM-SB-L", round_up(SupportLegs * SuppSmalls), round(FrontSupport))
                
                Hypotenuse =  RafterLChosen - 1000
                DistBetwFandB = Hypotenuse * np.cos(angle)
                ChangeElev = ((Hypotenuse)**2 - (DistBetwFandB)**2)**0.5
                BackSupport = FrontSupport + ChangeElev
                AddEntry("LM-SB-L", round_up(SupportLegs * SuppSmalls), round(BackSupport))
                
                BottomLen1 = DistBetwFandB/2
                CrossSupport1 = ((BottomLen1)**2 + (ChangeElev/2 + FrontSupport)**2)**0.5
                AddEntry("LM-SB-L", round_up(SupportLegs * SuppSmalls), round(CrossSupport1))
                
                BottomLen2 = DistBetwFandB - BottomLen1
                CrossSupport2 = ((BottomLen2)**2 + (ChangeElev/2 + FrontSupport)**2)**0.5
                AddEntry("LM-SB-L", round_up(SupportLegs * SuppSmalls), round(CrossSupport2))
                
        elif(pVert == 4 or pVert == 6):
            ATTC = SupportLegs
            AddEntry("LMK-ATTC", round_up(ATTC * ConSmalls), 0)
               
            #if(RafterLChosen >= 4200):
            SBConS = SupportLegs*2
            AddEntry("LM-TTC-90", round_up(SBConS * ConSmalls), 0)
            
            SBConD = SupportLegs*2
            AddEntry("LM-TTC-120", round_up(SBConD * ConSmalls), 0)
            
            FootPieceD = SupportLegs*2
            AddEntry("LM-FP-120", round_up(FootPieceD * ConSmalls), 0)
             
            FootPieceT = SupportLegs*1
            AddEntry("LM-FP-175", round_up(FootPieceT * ConSmalls), 0)
            
            FOHang =  (((pVert/2)*pLength+20*((pVert/2)-1)) - RafterLChosen)
            FrontSupport = (float(GroundClearanceE.get()) + FOHang*np.sin(angle))
            AddEntry("LM-SB-L", round_up(SupportLegs*2 * SuppSmalls), round(FrontSupport))
            
            Hypotenuse =  RafterLChosen - 500
            DistBetwFandB = Hypotenuse * np.cos(angle)
            ChangeElev = ((Hypotenuse)**2 - (DistBetwFandB)**2)**0.5
            BackSupport = FrontSupport + ChangeElev
            AddEntry("LM-SB-L", round_up(SupportLegs * SuppSmalls), round(BackSupport))
           
            BottomLen1 = DistBetwFandB/2
            CrossSupport1 = ((BottomLen1)**2 + (ChangeElev/2 + FrontSupport)**2)**0.5
            AddEntry("LM-SB-L", round_up(SupportLegs*2 * SuppSmalls), round(CrossSupport1))
                 
            BottomLen2 = DistBetwFandB - BottomLen1
            CrossSupport2 = ((BottomLen2)**2 + (ChangeElev/2 + FrontSupport)**2)**0.5
            AddEntry("LM-SB-L", round_up(SupportLegs*2 * SuppSmalls), round(CrossSupport2))
               
            
        # Cross Braces calculations:

        StandardBraceL = [2600, 3900, 4300] #LM-R20-L
        BraceL = 2600
        SupportSpaces = (SupportLegs/sysnum)-1
        NumberOfCrossSupport = (SupportSpaces//4 + 1)*sysnum # There should not be more than 4 spaces between supports. 5 is fine if only 2 are needed
        NumberOfCrossBraces = NumberOfCrossSupport*2

        global SupportSpacingL
        for i in range(len(StandardBraceL)):
            if (StandardBraceL[i] >= (SupportSpacingL + 300)):
                BraceL = StandardBraceL[i]
                break
        
        R20code = "LM-R20-"+str(BraceL)
        AddEntry(R20code, round_up(NumberOfCrossBraces * SuppSmalls), 0)
        
        SelfDrScr = 12*NumberOfCrossSupport
        AddEntry("FS-S-22X6-C4", round_up(SelfDrScr * SmallSmalls), 0)
        
        CapScr = round_up(NumberOfCrossSupport * SmallSmalls)
        AddEntry("FS-CAP-M8X60", CapScr, 0)
        
        FlatWashr = NumberOfCrossBraces
        AddEntry("FS-FW-M8", FlatWashr, 0)
       
        Nut = NumberOfCrossSupport
        AddEntry("FS-N-M8", Nut, 0)
        
        SpringWashr = NumberOfCrossSupport
        AddEntry("FS-SW-M8", SpringWashr, 0)
             
        if (Mount == 0): #Concrete Ground Mount

            ThreadRod = round_up(2*(FootPieceS + FootPieceD + FootPieceT) * SmallSmalls)
            AddEntry("FS-TR-M12X160", ThreadRod, 0)
             
            IKAnum = ((ThreadRod)//35 + 1)
            AddEntry("IKA-70007", IKAnum, 0)
            
            FlatW12 = ThreadRod*2
            AddEntry("FS-FW-M12", FlatW12, 0)
            
            Nut12 = ThreadRod*2
            AddEntry("FS-N-M12", Nut12, 0)
           
        elif (Mount == 1): #Ground Screw
            HB14x45 = round_up((FootPieceS + FootPieceD + FootPieceT)*2 * SmallSmalls)
            AddEntry("FS-HB-M14X45", HB14x45, 0)
             
            FlatW14 = 1*HB14x45
            AddEntry("FS-FW-M14", FlatW14, 0)
              
            SpringW14 = 1*HB14x45
            AddEntry("FS-SW-M14", SpringW14, 0)
            
            Nut14 = HB14x45
            AddEntry("FS-N-M14", Nut14, 0)
            
            GroundScrews = round_up((FootPieceS + FootPieceD + FootPieceT) * SuppSmalls)
            AddEntry("LM-GSF-ST-76X1600", GroundScrews, 0)
            

    elif (Mount == 2 and RafterLChosen < 4200): #Simple Pile

        GetPileOp()

        SBConS = 2*SupportLegs
        AddEntry("LM-TTC-60", round_up(SBConS * ConSmalls), 0)
        
        PConS = 2*SupportLegs
        AddEntry("LM-PC-60", round_up(PConS * ConSmalls), 0)
       
        # Pile Length calculations
        FOHang =  ((pVert*pLength+20*(pVert-1)) - RafterLChosen)/2
        FrontSupport = round((float(GroundClearanceE.get()) + FOHang*np.sin(angle) + EmbedmentD), 0)
        Hypotenuse =  RafterLChosen - 1000
        DistBetwFandB = Hypotenuse * np.cos(angle)
        ChangeElev = Hypotenuse*np.sin(angle)#((Hypotenuse)**2 - (DistBetwFandB)**2)**0.5
        BackSupport = round((FrontSupport + ChangeElev), 0)

        FPileQ = round_up(SupportLegs * SuppSmalls)
        FPileP = Pileprice(FrontSupport)
        FPileCode = PileCode
        FPileDescription = PileDescr.replace('specify', (str(FrontSupport)+"mm"), 1)
        FPentry = pd.DataFrame({"Code": [str(FPileCode)], 
                            "Description": [FPileDescription],
                            "Quantity": [FPileQ], 
                            "Price": [0],
                            "Discount": [str(discountp)+"%"],
                            "Discount Price": [FPileP],
                            "Total": [FPileP * FPileQ]})
        df = pd.concat([df, FPentry])
        
        BPileQ = round_up(SupportLegs * SuppSmalls)
        BPileP = Pileprice(BackSupport)
        BPileCode = PileCode.replace('F', 'R')
        BPileDescription = PileDescr.replace('FRONT', 'REAR', 1)
        BPileDescription = BPileDescription.replace('specify', (str(BackSupport)+'mm'), 1)
        BPentry = pd.DataFrame({"Code": [str(BPileCode)], 
                            "Description": [BPileDescription],
                            "Quantity": [BPileQ], 
                            "Price": [0],
                            "Discount": [str(discountp)+"%"],
                            "Discount Price": [BPileP],
                            "Total": [BPileP * BPileQ]})
        df = pd.concat([df, BPentry])
        
        #Cross-Bracing Calculations:
        
        StandardBraceL = [2600, 3900, 4300]
        BraceL = 2600
        SupportSpaces = (SupportLegs/sysnum)-1
        NumberOfCrossSupport = (SupportSpaces//4 + 1)*sysnum # There should not be more than 4 spaces between supports. 5 is fine if only 2 are needed
        NumberOfCrossBraces = round_up(NumberOfCrossSupport*2 * SuppSmalls)

        for i in range(len(StandardBraceL)):
            if (StandardBraceL[i] > (SupportSpacingL+300)):
                BraceL = StandardBraceL[i]
                break
            
        Bracewpm = 2.35
        CrossBracePr = round((Bracewpm * BraceL/1000 * Rate * MarkUp), 2) #2.35 is the kg/m weight
        
        crossbracecode = "LM-GM-CB-LC50"
        crossbracedescription = "GROUNDMOUNT CROSS BRACE DIM:50x25x10x2x3.2 LENGTH: *specify* PROFILE: LC FINNISH: PRE GALV MATERIAL: Z275"
        crossbracedescription = crossbracedescription.replace('specify', (str(BraceL)+'mm'), 1)
        
        CrossBrentry = pd.DataFrame({"Code": [crossbracecode], 
                            "Description": [crossbracedescription],
                            "Quantity": [NumberOfCrossBraces], 
                            "Price": [0],
                            "Discount": [str(discountp)+"%"],
                            "Discount Price": [CrossBracePr],
                            "Total": [CrossBracePr * NumberOfCrossBraces]})
        df = pd.concat([df, CrossBrentry])

        # Fasteners for the cross-bracing:
        M860Capnum = round_up(NumberOfCrossSupport * SmallSmalls)
        AddEntry("FS-CAP-M8X60", M860Capnum, 0)
        
        M835Capnum = round_up(4*NumberOfCrossSupport * SmallSmalls)
        AddEntry("FS-CAP-M8X35", M835Capnum, 0)
        
        M8FW = 2*M835Capnum + 2*M860Capnum
        AddEntry("FS-FW-M8", M8FW, 0)
        
        M8SW = 1*M835Capnum + 1*M860Capnum
        AddEntry("FS-SW-M8", M8SW, 0)

        M8N = 1*M835Capnum + 1*M860Capnum
        AddEntry("FS-N-M8", M8N, 0)
        
    elif (Mount == 3 or RafterLChosen >=4200): #Y Pile

        GetPileOp() 
        
        SBConS = 1*SupportLegs 
        AddEntry("LM-TTC-60", round_up(SBConS * ConSmalls), 0)

        SB2ConS = 2*SupportLegs
        AddEntry("LM-TTC-90", round_up(SB2ConS * ConSmalls), 0)

        P2ConS = 1*SupportLegs
        AddEntry("LM-PC-S-120", round_up(P2ConS * ConSmalls), 0)

        M12x110 = round_up(2*P2ConS * SmallSmalls)
        AddEntry("FS-HB-M12x110", M12x110, 0)

        M12x35 = round_up(2*P2ConS * SmallSmalls)
        AddEntry("FS-HB-M12X35", M12x35, 0)

        M12FW = 2*M12x35 + 2*M12x110
        AddEntry("FS-FW-M12", M12FW, 0)

        M12SW = 1*M12x35 + 1*M12x110
        AddEntry("FS-SW-M12", M12SW, 0)

        M12N = 1*M12x35 + 1*M12x110
        AddEntry("FS-N-M12", M12N, 0)

        PConS = round_up(1*SupportLegs * ConSmalls)
        AddEntry("LM-PC-60", PConS, 0)

        FOHang =  ((pVert*pLength+20*(pVert-1)) - RafterLChosen)/2
        FrontSupport = round((float(GroundClearanceE.get()) + FOHang*np.sin(angle) + EmbedmentD), 0)
        Hypotenuse =  RafterLChosen - 1000
        DistBetwFandB = Hypotenuse * np.cos(angle)
        ChangeElev = Hypotenuse*np.sin(angle)#((Hypotenuse)**2 - (DistBetwFandB)**2)**0.5
        BackSupport = round((FrontSupport), 0)

        FPileQ = round_up(SupportLegs * SuppSmalls)
        FPileP = Pileprice(FrontSupport)
        FPileCode = PileCode
        FPileDescription = PileDescr.replace('specify', (str(FrontSupport)+"mm"), 1)
        FPentry = pd.DataFrame({"Code": [str(FPileCode)], 
                            "Description": [FPileDescription],
                            "Quantity": [FPileQ], 
                            "Price": [0],
                            "Discount": [str(discountp)+"%"],
                            "Discount Price": [FPileP],
                            "Total": [FPileP * FPileQ]})
        df = pd.concat([df, FPentry])
        
        BPileQ = round_up(SupportLegs * SuppSmalls)
        BPileP = Pileprice(BackSupport)
        BPileCode = PileCode.replace('F', 'R')
        BPileDescription = PileDescr.replace('FRONT', 'REAR', 1)
        BPileDescription = BPileDescription.replace('specify', (str(BackSupport)+'mm'), 1)
        BPentry = pd.DataFrame({"Code": [str(BPileCode)], 
                            "Description": [BPileDescription],
                            "Quantity": [BPileQ], 
                            "Price": [0],
                            "Discount": [str(discountp)+"%"],
                            "Discount Price": [BPileP],
                            "Total": [BPileP * BPileQ]})
        df = pd.concat([df, BPentry])

        #Support Bar calcs:
        BPilePos = DistBetwFandB * 0.2
        BSuppBarL = round((ChangeElev**2 + BPilePos**2)**0.5)

        DistBetwFPandBP = DistBetwFandB * 0.8
        DistBetwFPandFSupp = Hypotenuse*np.cos(angle)
        DistBetwBPandFSupp = DistBetwFPandBP - DistBetwFPandFSupp
        FSuppBarL = round(((DistBetwBPandFSupp**2 + (0.6*ChangeElev)**2)**0.5))

        AddEntry("LM-SB-L", round_up(SupportLegs * SuppSmalls), BSuppBarL)
        
        AddEntry("LM-SB-L", round_up(SupportLegs * SuppSmalls), FSuppBarL)
        
        #Cross-Bracing Calculation:
        BraceL = round(((SupportSpacingL**2 + FrontSupport**2)**0.5), 0) #Small C-lipped channel
        SupportSpaces = (SupportLegs/sysnum)-1
        NumberOfCrossSupport = (SupportSpaces//4 + 1)*sysnum # There should not be more than 4 spaces between supports. 5 is fine if only 2 are needed
        NumberOfCrossBraces = round_up(NumberOfCrossSupport*2 * SuppSmalls)
        Bracewpm = 2.35
        CrossBracePr = round((Bracewpm * BraceL/1000 * Rate * MarkUp),2) #2.35 is the kg/m weight
        
        crossbracecode = "LM-GM-CB-LC50"
        crossbracedescription = "GROUNDMOUNT CROSS BRACE DIM:50x25x10x2x3.2 LENGTH: *specify* PROFILE: LC FINNISH: PRE GALV MATERIAL: Z275"
        crossbracedescription = crossbracedescription.replace('specify', (str(BraceL)+'mm'), 1)
        
        CrossBrentry = pd.DataFrame({"Code": [crossbracecode], 
                            "Description": [crossbracedescription],
                            "Quantity": [NumberOfCrossBraces], 
                            "Price": [0],
                            "Discount": [str(discountp)+"%"],
                            "Discount Price": [CrossBracePr],
                            "Total": [CrossBracePr * NumberOfCrossBraces]})
        df = pd.concat([df, CrossBrentry])

        # Fasteners for the cross-bracing:
        M860Capnum = round_up(NumberOfCrossSupport * SmallSmalls)
        AddEntry("FS-CAP-M8X60", M860Capnum, 0)

        M835Capnum = round_up(4*NumberOfCrossSupport * SmallSmalls)
        AddEntry("FS-CAP-M8X35", M835Capnum, 0)

        M8FW = 2*M835Capnum + 2*M860Capnum
        AddEntry("FS-FW-M8", M8FW, 0)

        M8SW = 1*M835Capnum + 1*M860Capnum
        AddEntry("FS-SW-M8", M8SW, 0)

        M8N = 1*M835Capnum + 1*M860Capnum
        AddEntry("FS-N-M8", M8N, 0)

def replace_first_l_with_numbers(input_str, replacement_numbers):
    count = 0
    result = ''

    for char in input_str:
        if char == 'L':
            count += 1
            if count == 1:
                result += str(replacement_numbers)  # Replace 'L' with the desired numbers
            else:
                result += char
        else:
            result += char

    return result

def getprice(code, quantity, length):
    global pricedf
    global discountp
    global MarkUpe
    global MarkUp
    global RafterLChosen
    
    discountp = float(DiscountE.get())

    discount = 1 - discountp/100
    
    MU = MarkUpE.get()

    if (MU is None or MU == 0):
        MarkUp = 1
    else:
        MarkUp = float(MU)/100 + 1

    ref = pricedf.iloc[:,0]
    prices = pricedf.iloc[:,2]
    descriptions = pricedf.iloc[:, 1]
    
    string = code

    index = 0

    if (code == "LM-R110-4200"):
        price = 1214.36 
        RafterLChosen = 4200
        description = "Rafter 110x4200mm AL6005 T6 Mill"
    else:
        for i in range(2,len(ref)):
            if (ref[i] == string):
                index = i
                
        price = prices[index]
        description = descriptions[index]
    
    if(code == "LM-SB-L"):
        pricet = (price*(length/1000)+17)
        description = replace_first_l_with_numbers(description, str(length))
    else:
        pricet = price

    price = round((pricet), 2)
    discprice = round(pricet*discount*MarkUp, 2)   
    totalprice = discprice*quantity
     
    
    return description, price, discprice, totalprice

def extract_length(s: str) -> int:
    match = re.search(r'\d+x\d+x(\d+)mm', s)
    if match:
        return int(match.group(1))
    raise ValueError("Length not found in string")    
    
def getStdSupportBarLength(Description):
    
    length = extract_length(Description)
    
    SBLengthArray = [450, 530, 550, 590, 615, 1500, 1550, 1560, 1740, 1800, 1815, 1870, 1875, 1960, 2525, 2540, 2610, 2615, 2670, 2710, 2795, 2820, 2840, 2890, 3000, 3340, 5000, 6000]
    
    if (length <= SBLengthArray[0]):
        return SBLengthArray[0]
    
    else:
        for i in range(1, len(SBLengthArray)):
            if length <= SBLengthArray[i] and length > SBLengthArray[i-1]:
                return SBLengthArray[i]
    match = re.search(r'\d+x\d+x(\d+)mm', s)
    if match:
        return int(match.group(1))
    raise ValueError("Length not found in string")
                
def AddK8Entry(code, quantity):
    global K8df
    
    NewEntry = pd.DataFrame({"Code": [code],
                             "Quantity": [str(quantity)]})
    K8df = pd.concat([K8df, NewEntry])

def ConvertToK8():
    global K8df
    
    K8df = pd.DataFrame(columns=['Code', 'Quantity'])
    
    K8Convertdf = pd.read_excel("Old and New Codes.xlsx")

    Oldcodedf = K8Convertdf.loc[:, ['Old Code']]
    Oldcodedf = Oldcodedf.dropna()
    OldcodeList = Oldcodedf.iloc[:,0].tolist()
    
    Newcodedf = K8Convertdf.loc[:, ['New Code']]
    Newcodedf = Newcodedf.dropna()
    NewcodeList = Newcodedf.iloc[:,0].tolist()
    
    QuoteCodesdf = df.loc[:, ['Code']]
    QuoteCodesdf = QuoteCodesdf.dropna()
    QuoteCodes = QuoteCodesdf.iloc[:,0].tolist()
    
    QuoteDescdf = df.loc[:, ['Description']]
    QuoteDescdf = QuoteDescdf.dropna()
    QuoteDescs = QuoteDescdf.iloc[:,0].tolist() 
    
    QuoteQuantitiesdf = df.loc[:, ['Quantity']]
    QuoteQuantitiesdf = QuoteQuantitiesdf.dropna()
    QuoteQuantities = QuoteQuantitiesdf.iloc[:,0].tolist()
    
    for i in range(0, len(QuoteCodes) - 1):
        for j in range(0, len(OldcodeList) - 1):
            
            if (QuoteCodes[i] == "LM-SB-L"):
                print(QuoteCodes[i])
                print(QuoteDescs[i])
                length = getStdSupportBarLength(QuoteDescs[i])
                StdSuppBarCode = "LM-SB-" + str(length)
                QuoteCodes[i] = StdSuppBarCode
                     
            if (QuoteCodes[i] == OldcodeList[j]):
                
                AddK8Entry(NewcodeList[j], QuoteQuantities[i])

def LoadWeights():
    global Weightdf
    
    Weightdf = pd.read_excel("Inventory Volume & weight.xlsx")
    Weightdf = Weightdf.iloc[8:, 0:3].reset_index(drop=True)
    Weightdf.columns = Weightdf.iloc[0]
    Weightdf = Weightdf.iloc[1:, :].reset_index(drop=True)
    Weightdf = Weightdf.iloc[:, [0, 2]]
    Weightdf = Weightdf.dropna()
    
    global WeightCode
    global Weights
    WeightCode = Weightdf.iloc[:,0].tolist()
    Weights = Weightdf.iloc[:,1].tolist()

def getWeight(code, description, quantity):
    global WeightCode
    global Weights
    
    for i in range(0, len(WeightCode) - 1):
        if (code == WeightCode[i]):
            if (code == "LM-SB-L"):
                Length = extract_length(description)/1000
                weight = round(Weights[i] * Length)
            else:
                weight = Weights[i]
            break
        elif ("LM-GM-P-F" in code):
            weight = round(Weightpm * FrontSupport)
        elif ("LM-GM-P-R" in code):
            weight = round(Weightpm * BackSupport)
        elif (code == "LM-GM-CB-LC50"):
            weight = round(Bracewpm * BraceL)
        else:
            weight = 0
    
    TotWeight = weight * int(quantity)
    TotWeight = round(float(TotWeight))
            
    return weight, TotWeight

def AddWeightEntry(weight, TotWeight):
    global quote_weight_df
    
    NewEntry = pd.DataFrame({"Unit Weight [g]": [float(weight)],
                             "Total Weight [g]": [float(TotWeight)]})
    quote_weight_df = pd.concat([quote_weight_df, NewEntry])

def CreateWeightDF():
    LoadWeights()
    
    global df
    global Weightdf
    
    # Create a new DataFrame with the same index as df
    global quote_weight_df
    quote_weight_df = pd.DataFrame(columns=['Unit Weight [g]', 'Total Weight [g]'])
    
    for i in range(1, len(df) - 1):
        code = df.iloc[i]['Code']
        description = df.iloc[i]['Description']
        quantity = df.iloc[i]['Quantity']
        
        weight, TotWeight = getWeight(code, description, quantity)
        
        # Add the weights to the new DataFrame
        AddWeightEntry(weight, TotWeight)
    
    #Adding total weight of the order    
    OrderWeight = (quote_weight_df['Total Weight [g]'].sum())/1000
    OrderWeight = round(float(OrderWeight), 3)
    OrderWeight = str(OrderWeight) + " kg"
    #AddWeightEntry("Total weight of the order", OrderWeight)
    NewEntry = pd.DataFrame({"Unit Weight [g]": ["Total weight of the order"],
                             "Total Weight [g]": [str(OrderWeight)]})
    quote_weight_df = pd.concat([NewEntry, quote_weight_df])
    
def CombineDataFrames(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    global df
    global K8df
    
    # Combine the two DataFrames
    df1 = df1.reset_index(drop=True)
    df2 = df2.reset_index(drop=True)
    
    return pd.concat([df1, df2], axis=1, ignore_index=False)

def AddEntry(code, quantity, length):
    global df
    global discountp
    
    description, price, discprice, total = getprice(code, quantity, length)
    
    NewEntry = pd.DataFrame({"Code": [code], 
                            "Description": [str(description)],
                            "Quantity": [quantity], 
                            "Price": [price],
                            "Discount": [str(discountp)+"%"],
                            "Discount Price": [discprice],
                            "Total": [total]})
    df = pd.concat([df, NewEntry])
      
def Calculations():
    debug = "Lol"
    debugLabel.config(text = debug)
    getPurlins()
    
def getDescription():
    global df
    global sysnum
    global angle
    global GroundClearance
    global pHor
    global pVert
    global pLength
    global pWidth
    global PurlinLMin
    global SupportLegsC
    global SupportSpacingL
    global OHang
    global EmbedmentD
    global Rate
    global MarkUp
    
    MountS = MountVar.get()
    
    if(MountS == 'Simple Piles' or MountS == 'Y-structure Piles'):
        description = MountS + " ground mount structure ("+str(pVert)+"V"+str(pHor)+ ") at " + str(round(np.rad2deg(angle))) + " degrees with " + str(GroundClearance) +" mm ground clearance for panels ("
        description = description + str(pWidth)+"x"+str(pLength)+" mm). Total length: "+str(PurlinLMin)+" mm. Number of Supports per structure: " + str(SupportLegsC)+". "
        description = description + "Support spacing: " + str(SupportSpacingL)+"mm with a purlin overhang of "+str(OHang)+"mm. "
        description = description + "Embedment Depth: "+str(EmbedmentD)+" mm."
        
    else:
        description = MountS + " ground mount structure ("+str(pVert)+"V"+str(pHor)+ ") at " + str(round(np.rad2deg(angle))) + " degrees with " + str(GroundClearance) +" mm ground clearance for panels ("
        description = description + str(pWidth)+"x"+str(pLength)+" mm). Total length: "+str(PurlinLMin)+" mm. Number of Supports per structure: " + str(SupportLegsC)+". "
        description = description + "Support spacing: " + str(SupportSpacingL)+"mm with a purlin overhang of "+str(OHang)+"mm. "
    
    total = round((df['Total'].sum()), 2)
    
    Descrentry = pd.DataFrame({"Code": ["DESCRIPTION"], 
                            "Description": [description],
                            "Quantity": [sysnum], 
                            "Price": [" "],
                            "Discount": [str(discountp)+"%"],
                            "Discount Price": [" "],
                            "Total": [total]})
    df = pd.concat([Descrentry, df])
    
def FinishCalc():
    
    getRaftChoice()
    MountSupp()
    getDescription()

    selection = "Success"
    debugLabel.config(text = selection)
    
def updateListBox(data):
    # clear list box
    ClientListBox.delete(0, END)
    
    # Add Clients to list box
    for item in data:
        ClientListBox.insert(END, item)
        
#Update entry box with listbox clicked
def fillout(e):
    #delete whatever is in the entry box
    CCodeE.delete(0, END)
    
    # Add clicked list item to entry box
    CCodeE.insert(0, ClientListBox.get(ACTIVE))
    
# Create function to check entry vs listbox
def check(e):
    # grab what was typed
    typed = CCodeE.get()
    
    if typed =='':
        data = CList
        updateListBox(data)
    else:
        data = []
        for item in CList:
            if typed.lower() in item.lower():
                data.append(item)
    
    updateListBox(data) 

def ProjectInfo():
    # Toplevel object which will 
    # be treated as a new window
    global newWindow
    newWindow = Toplevel(root)
 
    # sets the title of the
    # Toplevel widget
    newWindow.title("Project Information Entry")
 
    # sets the geometry of toplevel
    newWindow.geometry("1380x750")
    
    #Customer List collect
    CustomerListFrame = tk.LabelFrame(newWindow, text = "Load the latest customer list")
    CustomerListFrame.grid(row = 0, column = 0, padx = 5, pady = 5)
    
    #Customer Information
    CCodeLabel = tk.Label(CustomerListFrame, text = "Customer Details:")
    CCodeLabel.grid(row = 1, column = 0, padx = 5, pady = 5)
    global CCodeE
    CCodeE = tk.Entry(CustomerListFrame, width = 75)
    CCodeE.grid(row = 1, column = 1, padx = 5, pady = 5)
    global ClientListBox
    ClientListBox = tk.Listbox(CustomerListFrame, width = 75)
    ClientListBox.grid(row = 2, column = 1, padx = 5, pady = 5)
    
    # Create a binding on the listbox on click
    ClientListBox.bind("<<ListboxSelect>>", fillout)
    
    # Create a binding on the entry box
    CCodeE.bind("<KeyRelease>", check)
    
    #Button to find customer file
    CustomerListB = tk.Button(CustomerListFrame, text = "Load Customer List", command = lambda: Load_Customer_excel_data())
    CustomerListB.grid(row = 0, column = 0, padx = 5, pady = 5)
    
    #Project Details Entries
    PDFrame  = tk.LabelFrame(newWindow, text = "Enter the project details")
    PDFrame.grid(row = 0, column = 1, padx = 5, pady = 5)
    
    #Date
    DateLabel = tk.Label(PDFrame, text = "Please enter today's date (YYYY/MM/DD):")
    DateLabel.grid(row = 0, column = 0, padx = 5, pady = 5)
    global DateE
    DateE = tk.Entry(PDFrame)
    DateE.grid(row = 0, column = 1, padx = 5, pady = 5)
    
    #Reference
    ReferenceLabel = tk.Label(PDFrame, text = "Enter Quote Reference:")
    ReferenceLabel.grid(row = 1, column = 0, padx = 5, pady = 5)
    global ReferenceE
    ReferenceE = tk.Entry(PDFrame, width = 75)
    ReferenceE.grid(row = 1, column = 1, padx = 5, pady = 5)
    
    #Message
    MessageLabel = tk.Label(PDFrame, text = "Enter Quote Message:")
    MessageLabel.grid(row = 2, column = 0)
    global MessageE
    MessageE = tk.Entry(PDFrame, width = 75)
    MessageE.grid(row = 2, column = 1, padx = 5, pady = 5)

    #Buttons
    ButtonFrame  = tk.LabelFrame(newWindow, text = "Capture Information")
    ButtonFrame.grid(row = 3, column = 0, padx = 5, pady = 5)
    
    PIButton = tk.Button(ButtonFrame, text = "Capture Project Information", command = lambda: getProjectInfo())
    PIButton.grid(row = 0, column = 0, padx = 5, pady = 5)
    
def getProjectInfo():
    
    global transaction
    transaction = 'Quote'
    
    global date
    date = str(DateE.get())
    
    global QuoteRef
    QuoteRef = ReferenceE.get()
    
    global QuoteMessage
    QuoteMessage  = MessageE.get()
    
    global CustomerCode
    CCode = CCodeE.get()
    Customerarray = CCode.split('-')
    CustomerCode = Customerarray[0]
    
    global termname
    termname = 'CASH'
    
    global state
    state = 'Pending'
    
    global WarehouseID
    WarehouseID = 'Lumax - Olifantsfontein'
    
    global Unit
    Unit = 'Each'
    
    global DepartmentID
    DepartmentID = 'GroundMounting'
    
    global Sodcust
    Sodcust = 'CASH'
    
    newWindow.destroy()
    
def CreateSageImport():
    global df
    
    # Step 1: Read the quote template CSV file into a DataFrame
    template_file = 'import template.csv'
    template_df = pd.read_csv(template_file)

    # Display the template DataFrame
    #print("Template DataFrame:")
    #print(template_df.head())

    # Step 2: Assume you have a quote DataFrame with new data
    # Example quote DataFrame (replace this with your actual quote DataFrame)
    #quote_file = 'test.xlsx'
    #quote_df = pd.read_excel(quote_file)
    quote_df = df
    quote_df = quote_df.loc[:, ['Code', 'Description', 'Quantity', 'Discount Price']]
    quote_df.rename(columns={'Code': 'ITEMID', 'Description': 'ITEMDESC', 'Quantity': 'QUANTITY', 'Discount Price': 'PRICE'}, inplace=True)
                                            
    # Display the quote DataFrame
    #print("\nQuote DataFrame:")
    #print(quote_df.head())

    # Step 3: Create a list of dictionaries to represent rows for the merged DataFrame
    merged_data = []

    # Add the first row of the template DataFrame as the header row in the merged DataFrame
    merged_data.append(dict(zip(template_df.columns, template_df.iloc[0])))

    # Iterate over each row in the quote DataFrame and map it to the template columns
    for _, quote_row in quote_df.iterrows():
        # Create a dictionary to hold data for the new row
        new_row = {}

        # Map the quote data to the corresponding template columns
        for col in quote_df.columns:
            if col in template_df.columns:
                new_row[col] = quote_row[col]  # Assign quote data to the corresponding template column

        # Append the new row dictionary to the list
        merged_data.append(new_row)

    # Create the merged DataFrame directly from the list of dictionaries
    global merged_df
    merged_df = pd.DataFrame(merged_data, columns=template_df.columns)

    #Updating the line 1 items:
    merged_df.at[1, 'TRANSACTIONTYPE'] = transaction

    merged_df.at[1, 'DATE'] = date

    merged_df.at[1, 'GLPOSTINGDATE'] = date

    merged_df.at[1, 'CUSTOMER_ID'] = CustomerCode

    merged_df.at[1, 'TERMNAME'] = termname

    merged_df.at[1, 'REFERENCENO'] = QuoteRef

    merged_df.at[1, 'MESSAGE'] = QuoteMessage

    merged_df.at[1, 'STATE'] = state

    for i in range(1, (len(merged_df.index) - 1)):
        merged_df.at[i, 'LINE'] = i
        merged_df.at[i, 'WAREHOUSEID'] = WarehouseID
        merged_df.at[i, 'UNIT'] = "Each"
        merged_df.at[i, 'DEPARTMENTID'] = DepartmentID
        merged_df.at[i, 'LOCATIONID'] = "100 - Lumax"
        merged_df.at[i, 'SODOCUMENTENTRY_CUSTOMERID'] = Sodcust
        
    # Display the merged DataFrame
    #print("\nMerged DataFrame:")
    #print(merged_df)

    # Step 4: Save the merged DataFrame to a new CSV file
    Save_CSV()

def Save_CSV():
    global merged_df
    file = filedialog.asksaveasfilename(defaultextension = ".csv")
    merged_df.to_csv(str(file), index=False)
    label_file.config(text = "File saved")
    

root = tk.Tk()
root.geometry("1380x750")
root.title("Ground Mount Quote Tool (Use at your own risk)")

InputFrame = tk.LabelFrame(root, text = "Table Data Entry: ")
InputFrame.pack(side = "top", fill = "x")
#InputFrame.place(height = 400, width = 1380)

DispFrame = tk.LabelFrame(root, text = "Calculated Quote: ")
DispFrame.pack(expand = True,fill = "both")
#DispFrame.place(height = 350, width = 1380, rely = 0.525, relx = 0)

LoadPricesB = tk.Button(InputFrame, text = "Load current Prices", command = lambda: Load_excel_data())
LoadPricesB.grid(row = 1, column = 1, padx = 5, pady = 5)

CustomerInfoB = tk.Button(InputFrame, text = "Enter Project Info", command = lambda: ProjectInfo())
CustomerInfoB.grid(row = 1, column = 3, padx = 5, pady = 5)

label_file = ttk.Label(InputFrame, text = "No file selected")
label_file.grid(row = 1, column = 2, padx = 5, pady = 5)

debugLabel = tk.Label(InputFrame, text = "Lol")
debugLabel.grid(row = 1, column = 5, padx = 5, pady = 5)

SupportLabel = tk.Label(InputFrame, text = "Please email kavish@lumaxenergy.com to report any bugs.")
SupportLabel.grid(row = 1, column = 6, padx = 5, pady = 5)

TableNumberLabel = tk.Label(master = InputFrame, text = "Number of Tables:")
TableNumberLabel.grid(row = 2, column = 1, padx = 5, pady = 5)
TableNumberE = tk.Entry(InputFrame)
TableNumberE.grid(row = 2, column = 2, padx = 5, pady = 5)

MountLabel = tk.Label(InputFrame, text = "Mounting Choice:")
MountLabel.grid(row = 2, column = 3, padx = 5, pady = 5)
MountVar = tk.StringVar()
MountStr = ['Concrete Ground Mount', 'Ground Screw', 'Simple Piles', 'Y-structure Piles']
MountVar.set(MountStr[0])
MountOp = tk.OptionMenu(InputFrame, MountVar, *MountStr)
MountOp.grid(row = 2, column = 4, padx = 5, pady = 5)


VertPan = ['1','2','3', '4', '6']
VertVar = tk.StringVar()
VertVar.set(VertPan[0])
VertPanelLabel = tk.Label(InputFrame, text = "Number of Vertical Panels (2 or 3):")
VertPanelLabel.grid(row = 4, column = 1, padx = 5, pady = 5)
VertPanelOp = tk.OptionMenu(InputFrame, VertVar, *VertPan)
VertPanelOp.grid(row = 4, column = 2, padx = 5, pady = 5)

SharedRailLabel = tk.Label(InputFrame, text = "Shared rails:")
SharedRailLabel.grid(row = 3, column = 1, padx = 5, pady = 5)
ShRailVar = tk.StringVar()
ShRailList = ['Yes', 'No']
ShRailVar.set(ShRailList[1])
ShRailOp = tk.OptionMenu(InputFrame, ShRailVar, *ShRailList)
ShRailOp.grid(row = 3, column = 2, padx=5, pady=5)

HorPanelLabel = tk.Label(InputFrame, text = "Number of Horizontal panels per row:")
HorPanelLabel.grid(row = 5, column = 1, padx = 5, pady = 5)
HorPanelE = tk.Entry(InputFrame)
HorPanelE.grid(row = 5, column = 2, padx = 5, pady = 5)

DiscountLabel = tk.Label(InputFrame, text = "Customer Discount [%]:")
DiscountLabel.grid(row = 7, column = 1, padx = 5, pady = 5)
DiscountE = tk.Entry(InputFrame)
DiscountE.grid(row = 7, column = 2, padx = 5, pady = 5)

PanelWidthLabel = tk.Label(InputFrame, text = "Width of the selected panels:")
PanelWidthLabel.grid(row = 3, column = 3, padx = 5, pady = 5)
PanelWidthE = tk.Entry(InputFrame)
PanelWidthE.grid(row = 3, column = 4, padx = 5, pady = 5)

PanelLengthLabel = tk.Label(InputFrame, text = "Length of the selected panels:")
PanelLengthLabel.grid(row = 4, column = 3, padx = 5, pady = 5)
PanelLengthE = tk.Entry(InputFrame)
PanelLengthE.grid(row = 4, column = 4, padx = 5, pady = 5)

AngleLabel = tk.Label(InputFrame, text = "Angle (degrees):")
AngleLabel.grid(row = 5, column = 3, padx = 5, pady = 5)
AngleE = tk.Entry(InputFrame)
AngleE.grid(row = 5, column = 4, padx = 5, pady = 5)

GroundClearanceLabel = tk.Label(InputFrame, text = "Ground Clearance:")
GroundClearanceLabel.grid(row = 6, column = 3, padx = 5, pady = 5)
GroundClearanceE = tk.Entry(InputFrame)
GroundClearanceE.grid(row = 6, column = 4, padx = 5, pady = 5)

MaxSuppSpaceLabel = tk.Label(InputFrame, text = "Maximum Space Between Supports:")
MaxSuppSpaceLabel.grid(row = 7, column = 3, padx = 5, pady = 5)
MaxSuppSpaceE = tk.Entry(InputFrame)
MaxSuppSpaceE.grid(row = 7, column = 4, padx = 5, pady = 5)

EmbedmentDLabel = tk.Label(InputFrame, text = "Embedment Depth:")
EmbedmentDLabel.grid(row = 3, column = 5, padx = 5 , pady = 5)
EmbedmentDE = tk.Entry(InputFrame)
EmbedmentDE.grid(row = 3, column = 6, padx = 5, pady = 5)

MemberList()
PileSLabel = tk.Label(InputFrame, text = "Pile size:")
PileSLabel.grid(row = 4, column = 5, padx = 5, pady = 5)
PileSVar = tk.StringVar()
#PileSList = ['150x75x20x3 mm', '150x70x15x3 mm', '125x65x20x3 mm']
PileSList = Piledf['Pile Description'].tolist()
PileSVar.set(PileShortList[0])
PileSOp = tk.OptionMenu(InputFrame, PileSVar, *PileShortList)
PileSOp.grid(row = 4, column = 6, padx=5, pady=5)

RateLabel = tk.Label(InputFrame, text = "Rate (R/kg):") 
RateLabel.grid(row = 5, column = 5, padx = 5, pady = 5)
RateE = tk.Entry(InputFrame)
RateE.grid(row = 5, column = 6, padx = 5, pady = 5)

MarkUpLabel = tk.Label(InputFrame, text = "Markup (%)")
MarkUpLabel.grid(row = 6, column = 5, padx = 5, pady =5)
MarkUpE = tk.Entry(InputFrame)
MarkUpE.grid(row = 6, column = 6, padx = 5, pady = 5)

ROHLabel = tk.Label(InputFrame, text = "Please select a total panel overhang on rafter:")
ROHLabel.grid(row = 6, column = 1, padx = 5, pady = 5)
var = tk.StringVar()
RaftOvList = ['600mm', '800mm']
var.set(RaftOvList[0])
RaftOvOp = tk.OptionMenu(InputFrame, var, *RaftOvList)
RaftOvOp.grid(row = 6, column = 2, padx = 5, pady = 5)

SSmallsLabel = tk.Label(InputFrame, text = "Extra Fasteners and Clamps Percentage:")
SSmallsLabel.grid(row = 8, column = 1, padx = 5, pady = 5)
global SSmallsVar
SSmallsVar = tk.StringVar()
SSmallsList = ['2%', '5%', '10%']
SSmallsVar.set(SSmallsList[0])
SSmallsOp = tk.OptionMenu(InputFrame, SSmallsVar, *SSmallsList)
SSmallsOp.grid(row = 8, column = 2, padx = 5, pady = 5)

ConSmallsLabel = tk.Label(InputFrame, text = "Extra Connectors(TTC's, FP's) Percentage:")
ConSmallsLabel.grid(row = 8, column = 3, padx = 5, pady = 5)
global ConSmallsVar
ConSmallsVar = tk.StringVar()
ConSmallsList = ['0%', '5%', '10%']
ConSmallsVar.set(ConSmallsList[0])
ConSmallsOp = tk.OptionMenu(InputFrame, ConSmallsVar, *ConSmallsList)
ConSmallsOp.grid(row = 8, column = 4, padx = 5, pady = 5)

SuppSmallsLabel = tk.Label(InputFrame, text = "Extra Supports Percentage:")
SuppSmallsLabel.grid(row = 8, column = 5, padx = 5, pady = 5)
global SuppSmallsVar
SuppSmallsVar = tk.StringVar()
SuppSmallsList = ['0%', '5%', '10%']
SuppSmallsVar.set(SuppSmallsList[0])
SuppSmallsOp = tk.OptionMenu(InputFrame, SuppSmallsVar, *SuppSmallsList)
SuppSmallsOp.grid(row = 8, column = 6, padx = 5, pady = 5)

CalcRaftB = tk.Button(InputFrame, text = "Calculate Rafter Length", command = lambda: Calculations())
CalcRaftB.grid(row = 9, column = 1, padx = 5, pady = 5)
CalcRaftLabel = tk.Label(InputFrame, text = " ")
CalcRaftLabel.grid(row = 9, column = 2, padx = 5, pady = 5)

RafterChoiceLabel = tk.Label(InputFrame, text = "Please select one of the standard Rafter Lengths in mm:")
RafterChoiceLabel.grid(row = 10, column = 1, padx = 5, pady = 5)
RaftVar = tk.StringVar()
#RaftStr = ['3400', '3600', '3800', '4000', '4200', '4400', '5400', '5600', '6200']
RaftStr = Rafterdf['Rafter Description'].tolist()
RaftVar.set(RaftStr[0])
RafterChoiceOp = tk.OptionMenu(InputFrame, RaftVar, *RaftStr)
RafterChoiceOp.grid(row = 10, column = 2, padx = 5, pady = 5)

CalcPurlLabel = tk.Label(InputFrame, text = "Calculated Purlin Length")
CalcPurlLabel.grid(row = 11, column = 1, padx = 5, pady = 5)

PurlinLabel = tk.Label(InputFrame, text = "Supplied Purlin Length:")
PurlinLabel.grid(row = 11, column = 2, padx = 5, pady = 5)

SupportSLabel = tk.Label(InputFrame, text = "Support Spacing")
SupportSLabel.grid(row = 11, column = 3, padx = 5, pady = 5)

SupportLegsLabel = tk.Label(InputFrame, text = "Support Legs")
SupportLegsLabel.grid(row = 11, column = 4, padx = 5, pady = 5)

OHangLabel = tk.Label(InputFrame, text = "Overhang")
OHangLabel.grid(row = 11, column = 5, padx = 5, pady = 5)

TotalPriceLabel = tk.Label(InputFrame, text = "Total Price of the quote:")
TotalPriceLabel.grid(row = 11, column = 6, padx = 5, pady = 5)

CalcQuoteB = tk.Button(InputFrame, text = "Calculate Quote", command = lambda: FinishCalc())
CalcQuoteB.grid(row = 12, column = 1, padx = 5, pady = 5)

DispQuoteB = tk.Button(InputFrame, text = "Display Quote", command = lambda: Refresh())
DispQuoteB.grid(row = 12, column = 2, padx = 5, pady = 5)

ExportB = tk.Button(InputFrame, text = "Export Quote", command = lambda: Save_Excel())
ExportB.grid(row = 12, column = 3, padx = 5, pady = 5)

SageB = tk.Button(InputFrame, text = "Create Sage Import", command = lambda: CreateSageImport())
SageB.grid(row = 12, column = 4, padx = 5, pady = 5)

# Treeview Widget
tv1 = ttk.Treeview(DispFrame)
tv1.place(relheight=1, relwidth=1)

treescrolly = tk.Scrollbar(DispFrame, orient = "vertical", command=tv1.yview)
treescrollx = tk.Scrollbar(DispFrame, orient = "horizontal", command = tv1.xview)
tv1.configure(xscrollcommand = treescrollx.set, yscrollcommand = treescrolly.set)
treescrollx.pack(side = "bottom", fill = "x")
treescrolly.pack(side = "right", fill = "y")

# Add weights to the grid rows and columns
# Changing the weights will change the size of the rows/columns relative to each other
DispFrame.grid_rowconfigure(0, weight=1)
DispFrame.grid_rowconfigure(1, weight=1)
DispFrame.grid_columnconfigure(0, weight=1)
DispFrame.grid_columnconfigure(1, weight=1)

root.mainloop()