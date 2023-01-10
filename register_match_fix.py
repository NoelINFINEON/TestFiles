#NOTE: Trimming Functions found in this File 
from os import write
# import pytest
# import pylink
# import pyvisa
from PyPDF2 import PdfFileMerger #use command ---->pip install PyPDF2 to install PDF merger file 
import time
from importRegMap import * #imports from another file
# from test_noel import *
from smu_commands import *
# from test_excel import *
from connections import *

# from test_excel import *
# from smu_commands import turn_off_smu
# from start_helper_functions import helper_functions
# db = helper_functions()

# SmuKeysight =  pyvisa.ResourceManager().open_resource('GPIB0::24::INSTR') #connect to the SMU
# powerSupply =  pyvisa.ResourceManager().open_resource('GPIB0::08::INSTR') #connect to Power Supply 
# dataAcq =  pyvisa.ResourceManager().open_resource('GPIB0::10::INSTR') #connect to the data Acq Unit



# global jlink
# jlink = pylink.JLink()
# jlink.open()
# jlink.set_tif(pylink.enums.JLinkInterfaces.SWD)
# jlink.connect('CORTEX-M0')
"""
def write_specific_bits(x, start_bit, field_size, val, word_size=8):
    #NOTE: Figure out a way to read what bits need to be changed throught the PFD somehow, 
    #Try doing as follows: READ REGISTER -> Save into variable, Modify with New Value -> Write Back Into Reg
    # x is original binary value, start bit is start of bit position, field size is the number of bits youd like to change, val is value that
    #will replace those bits, word size is the size of bits original x value
    #link that explains this--> https://stackoverflow.com/questions/65622010/how-to-modify-multiple-bits-in-an-integer         
    mask = 2 ** word_size - 1 - (2 ** field_size - 1) * (2 ** start_bit)
    val = min(val, 2 ** (field_size) - 1)
    w = val * 2 ** start_bit
    y = (x & mask) | w
    return (y)
"""

"""BASIC BLOCKS FUNCTIONS"""
#May be edited to do all basic block trims in here 
def SweepSearchBG(jlink,debugging,reg_name,target_voltage):
    #Array Initiliazation
    trimlist = [] #define the array 
    trimCode = []
    trimReg = []
    dataList1 = []
    regWrite = []
    lowestValue = []
    closestNum = []
    muxA = []
    a_list = [0]
    b_list = list(range(32,33))
    #  print ("Hello!")
    #  print (trimAddy)

    #trimRegString = "0x70000000"
    #trimRegHex = int(trimAddy,16)
    trimData = "0x00000000"
    trimDataHex = int(trimData,16)
    dataList1.append(trimDataHex)
    #  jlink.open()
    data = [dataList1[0]]
    # jlink.memory_write32(trimAddy, dataList1)
    findREG(db.jlink,debugging,reg_name, dataList1)

    #################################################################################################################

    #This will fill up an array with the BG readings from the DAQ. - WORKING
    n = 64
    b = 0
    A = db.get_global("BG_DAQ") #setting up the channel number
    # B = db.get_global("BG_DAQ")
    DAQ_ch1 = db.get_global("CH1_DAQ")  #pad_avrrdy
    DAQ_ch2 = db.get_global("CH2_DAQ")  #pad_avren
    print("---------- Sweep search START ----------")
    for i in range(0, n): 
        # jlink.memory_write32(trimAddy, a_list) #note that a_list only has one element at all times
        findREG(db.jlink,debugging,reg_name, a_list)
        # dataAcq.write("CONFIGURE:VOLTAGE:DC AUTO, (@"+A+")")
        # dataAcq.write("SENSE:VOLT:DC:NPLC 2,"+"(@"+A+")")
        # dataAcq.write("INPUT:IMPEDANCE:AUTO ON,(@"+A+")")
        # dataAcq.write("ZERO:AUTO ON,(@"+A+")")
        # dataAcq.write("READ?")
        b = b + 1
        a_list[0] = b
        # BG_value_test = dataAcq.read_ascii_values()
        ele = abs(DAQread(DAQ_ch1)-DAQread(DAQ_ch2))
        # BG_value_true = BG_value_test[0]
        # ele = BG_value_true
        trimlist.append(ele) # adding whatever the DAQ reads into an array and appeneds each reading
        # print(trimlist if len(trimlist)>1)
        # db.printing(debugging,trimlist)
        print("code= ", b-1,", read_val=  ", ele)
    #################################################################################################################
    bg_desired = target_voltage # set desired trim reading value here
    n = 64
    b = 0
    z = 0
    for i in range(0,n):
        bg_measured = trimlist[z]
        b_list[0] = z 
        Perc_Diff = (abs(bg_desired-bg_measured))/((bg_desired+bg_measured)/2)*100 #calculates measured vs desired
        if Perc_Diff<.1: #stores all trim values that fall within the percent threshold in an array
            #jlink.memory_write32(trimRegHex,b_list)
            lowestValue.append(b_list[0])
            muxA.append(bg_measured)
            # print (lowestValue)
    #write trimBGcode into the register here 
        z = z+1
    #######################################################################################################3
    #  print (trimlist)
    #  print (muxA)
    print ("\n\n**-> Desired Trim Value:",bg_desired)
    closestReading = muxA[min(range(len(muxA)), key = lambda z: abs(muxA[z]-bg_desired))]
    location = muxA.index(closestReading)
    FINALLY = lowestValue[location]
    a_list[0] = FINALLY
    print ("**-> Your Trim Value is  :", FINALLY)
    print ("**-> Your closest reading is : ", closestReading)
    print ("-"*50)
    adressLocc =1879048192
    swd_argzzz = []
    swd_argzzz.append(FINALLY)
    data = [swd_argzzz[0]]
    # jlink.memory_write32(adressLocc, data)
    findREG(db.jlink,debugging,reg_name, data)
    db.store_global("BG_Final",closestReading)
    db.store_global("BG_trim_code",FINALLY)
    global BG_trim_value
    BG_trim_value = FINALLY

    return [FINALLY,closestReading]
    #can make edits above to improve 



    

    
    
    
    
    
    #note that a_list only has one element at all times
    #jlink.memory_write32(trimAddy,FINALLY)



    """
    def closest(lowestValue, bg_desired):#will find closest value to that of desired value and write into the register
        FinalTrimCode = lowestValue[min(range(len(lowestValue)), key = lambda i: abs(lowestValue[i]-bg_measured))]
        b_list[0] = FinalTrimCode
        jlink.memory_write32(trimAddy,b_list)
        return lowestValue[min(range(len(lowestValue)), key = lambda i: abs(lowestValue[i]-bg_measured))]
    print (closest(lowestValue, bg_desired))
    print (b_list)
    """
def Read_BG_volt(Board_addr):
    CH = db.get_global("BG_DAQ")
    dataAcq.write("CONFIGURE:VOLTAGE:DC AUTO, (@"+CH+")")
    dataAcq.write("SENSE:VOLT:DC:NPLC 2,(@"+CH+")")
    dataAcq.write("INPUT:IMPEDANCE:AUTO ON,(@"+CH+")")
    dataAcq.write("ZERO:AUTO ON,(@"+CH+")")
    dataAcq.write("READ?")
    bg_volt = dataAcq.read_ascii_values()[0]
    print("BG Volt = ",bg_volt)
    return bg_volt

def BinarySearchBG(jlink,debugging,reg_name,target_voltage):
    CH = db.get_global("BG_DAQ")

    low = 0
    high = 64
    mid = 0
    mid_val=0
    searching = 1
    bg_trim_code = 0
    bg_trim_val = 0
    DAQ_ch1 = db.get_global("CH1_DAQ")  #pad_avrrdy
    DAQ_ch2 = db.get_global("CH2_DAQ")  #pad_avren

    print("---------- Binary search START ----------")
    while (low <= high) and (searching ==1):
        mid = int((low+high)/2)  

        # jlink.memory_write32(trim_address, [mid]) #note that a_list only has one element at all times
        findREG(db.jlink,debugging,reg_name, [mid])
        # dataAcq.write("CONFIGURE:VOLTAGE:DC AUTO, (@"+CH+")")
        # dataAcq.write("SENSE:VOLT:DC:NPLC 2,(@"+CH+")")
        # dataAcq.write("INPUT:IMPEDANCE:AUTO ON,(@"+CH+")")
        # dataAcq.write("ZERO:AUTO ON,(@"+CH+")")
        # dataAcq.write("READ?")
        # mid_val = dataAcq.read_ascii_values()[0]
        mid_val = abs(DAQread(DAQ_ch1)-DAQread(DAQ_ch2))
        if mid_val < target_voltage:
            low = mid +1
        elif mid_val> target_voltage:
            high = mid -1
        elif mid_val == target_voltage:
            searching =0
        
        print("code= ", mid,", read_val=  ", mid_val)
        
        if searching == 0: # if target an exact trim code
            bg_trim_code = mid
            bg_trim_val = mid_val
    # print ("searching=",
    # searching)
    # print ("high:",high)
    if searching == 1: # if target not exact trim code, take nearest code
        if (mid_val > target_voltage):
            bg_trim_code = low
            
        else:
            bg_trim_code = high
            # print ("low:",low)
    print("---------- Binary search END ----------")

    findREG(db.jlink,debugging,reg_name, [bg_trim_code])
    # # jlink.memory_write32(trim_address, [bg_trim_code]) 
    # dataAcq.write("CONFIGURE:VOLTAGE:DC AUTO, (@"+CH+")")
    # dataAcq.write("SENSE:VOLT:DC:NPLC 2, (@"+CH+")")
    # dataAcq.write("INPUT:IMPEDANCE:AUTO ON,(@"+CH+")")
    # dataAcq.write("READ?")
    # bg_trim_val = dataAcq.read_ascii_values()[0]
    bg_trim_val = abs(DAQread(DAQ_ch1)-DAQread(DAQ_ch2))
    
    print ("**-> Desired Trim Value:",target_voltage)
    print ("**-> Your Trim code is  :", bg_trim_code)
    print ("**-> Your closest reading is : ", bg_trim_val)
    print ("-"*50)
    # adressLocc =1879048192 # bg trim reg 0x70000000
    # jlink.memory_write32(adressLocc, [bg_trim_code])
    db.store_global("BG_Final",bg_trim_val)
    db.store_global("BG_trim_code",bg_trim_code)
    # findREG(db.jlink,debugging,reg_name, [bg_trim_code])
    # global BG_trim_value
    # BG_trim_value = bg_trim_code     
    return [bg_trim_code,bg_trim_val]   

def ts_current_trim(db,jlink,debugging,trim_address,TS_current_target,method,reg_name_for_iterating_code):
    # read atsen current - channel 2 
    pad = 'pad_atsen'
    trim_val = 0.0
    ts_target = float(TS_current_target)*1e-6
    # print("-"*20,"TS Trim start","-"*20)
    if method == 0: # sweep
        curr = []
        err = []
        for i in range(0,63):
            db.printing(debugging,"-"*50)
            findREG(db.jlink,debugging, reg_name_for_iterating_code, i)
            # print(i)
            # b = read_curr_from_smu(pad)
            curr.append(read_curr_from_smu(pad))
            err.append(100)
            err[i] = abs(abs(curr[i][0])-ts_target)
            print("**-> code= ", i,", read_val=  ", abs(curr[i][0]))
        final = min(err)
        code = err.index(final)
        trim_val = curr[code]

    else:   # binary search
        low = 0
        high = 63
        mid = 0
        mid_val=0        
        searching = 1
        while (low <= high) and (searching ==1):
            mid = int((low+high)/2)  
            db.printing(debugging,"-"*50)
            findREG(db.jlink,debugging, reg_name_for_iterating_code, mid)
            # b = read_curr_from_smu(pad)
            # print(b)
            mid_val = abs(read_curr_from_smu(pad)[0])
            if mid_val < ts_target:
                low = mid +1
            elif mid_val> ts_target:
                high = mid -1
            elif mid_val == ts_target:
                searching =0
            if searching == 0: # if target an exact trim code
                code = mid
                trim_val = mid_val
            print("code= ", mid,", read_val=  ", mid_val)
        if searching == 1: # if target not exact trim code, take nearest code
            if (mid_val < ts_target):
                code = high
            else:
                code = low 
    if code >=64:
        code = 63        
    findREG(db.jlink,debugging, reg_name_for_iterating_code, code)
    trim_val = abs(read_curr_from_smu(pad)[0])
    # print(trim_val)
    print ("**-> Desired Trim Value:",TS_current_target,"uA")
    # print ("**-> Your Trim code is  :", code)
    # print ("**-> Your closest reading is : ", trim_val)
    findREG(db.jlink,debugging, reg_name_for_iterating_code, code)
    return [code,trim_val]

def VGPTAT_binary(jlink,debugging,Target_VPtat_temp,reg_name_for_iterating_code):
    DAQ_ch1 = db.get_global("CH1_DAQ")  #pad_avrrdy
    DAQ_ch2 = db.get_global("CH2_DAQ")  #pad_avren
    low = 0
    high = 63
    mid = 0
    mid_val=0        
    searching = 1
    while (low <= high) and (searching ==1):
        mid = int((low+high)/2)  
        db.printing(debugging,"-"*50)
        findREG(db.jlink,debugging, reg_name_for_iterating_code, mid)
        mid_val = abs(DAQread(DAQ_ch1)-DAQread(DAQ_ch2))
        if mid_val < Target_VPtat_temp:
            low = mid +1
        elif mid_val> Target_VPtat_temp:
            high = mid -1
        elif mid_val == Target_VPtat_temp:
            searching =0
        if searching == 0: # if target an exact trim code
            code = mid
            trim_val = mid_val
        print("code= ", mid,", read_val=  ", mid_val)
    if searching == 1: # if target not exact trim code, take nearest code
        if (mid_val > Target_VPtat_temp):
            code = high
        else:
            code = low 
    if code >=64:
        code = 63        
    findREG(jlink,debugging, reg_name_for_iterating_code, code)
    mid_val = abs(DAQread(DAQ_ch1)-DAQread(DAQ_ch2))
    # trim_val = abs(read_curr_from_smu(pad)[0])
    # print(trim_val)
    print ("**-> Desired Trim Value:",Target_VPtat_temp)
    print ("**-> Your Trim code is  :", code)
    print ("**-> Your closest reading is : ", mid_val)
    findREG(db.jlink,debugging, reg_name_for_iterating_code, code)
    return [code,mid_val]

def VGPTAT_trimMe(db,jlink,debugging,trimlist,daqReadVal, Target_VPtat_temp): #replace here with an SMU read since the DAQ is not used for trimming
    global FINALLY_VGPTAT
    lowestValue = []
    muxA = []
    finalVal = [0]
    # set desired trim reading value here
    daqReadVal.pop(0) #remove first element of array due to wrong register
    trimlist.pop(0) #remove first element if part is already TRIMMED
    for i in range(len(daqReadVal)):
        lo_measured = daqReadVal[i]
        Perc_Diff = (abs(Target_VPtat_temp-lo_measured))/((Target_VPtat_temp+lo_measured)/2)*100 #calculates measured vs desired
        if Perc_Diff<1: #stores all trim values that fall within the percent threshold in an array
            #jlink.memory_write32(trimRegHex,b_list)
            lowestValue.append(trimlist[i]) #filling array with trim codes
            muxA.append(lo_measured)
    truename = "trim.BG_TRIM1.bg_vptat_trim"
    #realList = removeUnwantedBitsFromTrimList(truename,lowestValue)

    closestReading = muxA[min(range(len(muxA)), key = lambda z: abs(muxA[z]-Target_VPtat_temp))]
    location = muxA.index(closestReading)
    FINALLY_VGPTAT = lowestValue[location]
    print ("**-> Your Desired Trim Val for VGPTAT:", Target_VPtat_temp)
    print ("**-> Your Trim code for VGPTAT is  :", FINALLY_VGPTAT)
    print ("**-> Your closest reading for VGPTAT is : ", closestReading)
    print ("-"*50)
    db.printing(debugging,"Your Dsired Trim Val for VGPTAT:", Target_VPtat_temp)
    db.printing(debugging,"Your Trim code for VGPTAT is  :", FINALLY_VGPTAT)
    db.printing(debugging,"Your closest reading for VGPTAT is : ", closestReading)
    db.printing(debugging,"-"*50)
    return [FINALLY_VGPTAT,closestReading]

#Code Update as of 01/08/2023
#Trim Algorith Funcitons for SP Block and IS_LR Condensed down to one function

def trim_sweep(db,jlink,debugging,trimAddy,trimValues,MUX_READ,target_voltage,name):
    lowestValue = [] #trim readings that fall within percent threshold stored here
    muxA = [] #DAQ readings that fall within perent threshhold stored here
    MUX_READ.pop(0) #remove first element of array due to wrong register
    trimValues.pop(0) #remove first element if part is already TRIMMED

    for i in range(len(MUX_READ)):
        lo_measured = MUX_READ[i]
        Perc_Diff = (abs(target_voltage-lo_measured))/((target_voltage+lo_measured)/2)*100 #calculates measured vs desired
        if Perc_Diff<1: #stores all trim values that fall within the percent threshold in an array
            lowestValue.append(trimValues[i]) #filling array with trim codes
            muxA.append(lo_measured)

    truename = name
    #removeUnwantedBits is used for SP block trim codes to remove zeros from trim values
    realList = removeUnwantedBitsFromTrimList(jlink,debugging,truename,lowestValue)

    #Algorithm Below will sweep through muxA array and find closest value to target voltage
    closestReading = muxA[min(range(len(muxA)), key = lambda z: abs(muxA[z]-target_voltage))]
    #Once detecting closest voltage, position of correlating trim value is stored into location
    location = muxA.index(closestReading)
    #Final Trim Value Is Stored to variable final_trim_code
    final_trim_code = realList[location]
    print ("**-> Your target Trim Val for -----:", target_voltage)
    print ("**-> Your Trim code for ----- is  :", final_trim_code)
    print ("**-> Your closest reading for ----- is : ", closestReading)
    print ("-"*50)
    db.printing(debugging,"Your target Trim Val for -----:", target_voltage)
    db.printing(debugging,"Your Trim code for ----- is  :", final_trim_code)
    db.printing(debugging,"Your closest reading for ----- is : ", closestReading)
    db.printing(debugging,"-"*50)


    return [final_trim_code,closestReading]





















#DAQRead will read the voltage where CH1 of the Daq is connected
def DAQread(CH_SEL): 

    A = CH_SEL #Setting Channel Here , can set more Channels as needed
    B = "2"
    dataAcq.write("CONFIGURE:VOLTAGE:DC AUTO, (@"+A+")")

    dataAcq.write("SENSE:VOLT:DC:NPLC "+B+","+"(@"+A+")")

    dataAcq.write("INPUT:IMPEDANCE:AUTO ON,(@"+A+")")

    dataAcq.write("ZERO:AUTO ON,(@"+A+")")

    dataAcq.write("READ?")
    daqREADING = dataAcq.read_ascii_values()
    realdaqREADING = daqREADING[0]
    ele1 = realdaqREADING
    # db.printing(debugging,ele1)
    return ele1



def read_temp_from_daq():
    thermocouples = db.get_global("thermocouples")
    daqREADING = 0
    for i in thermocouples:
        dataAcq.write("CONFIGURE:TEMP TC,K, (@"+str(i)+")")
        dataAcq.write("ROUT:SCAN (@"+str(i)+")")
        dataAcq.write("READ?")
        daqREADING += dataAcq.read_ascii_values()[0]
    final_temp = daqREADING/len(thermocouples)
    return (final_temp)    
    # print("Temperature= ",daqREADING[0],end='\r')




#Code below is in progress





 



#Dont need to add the calculation multiple times, equations are the same 
#add a way to know when to call this function 
# def ISResidueCalculations():
#     a = v_one_k #v_one_k is equal to 
#     b = pad_isen1 #equal to the current being applied 
#     r_meas=(a/b)
#     r_err= ((1.0e3 - r_meas)/r_meas)
#     is_500_ohm_res_tcode = ((r_err * 10.0) + 3.0)
#     is_one_kohm_res_tcode = ((r_err * 20.0) + 3.0)
#     #save each important variable in an array , and when writing
#     #res trim , look for these variables that are already stored 
#     return











