"""
This python script uses an edited version of Brady Pierce's "Stepper_Job.py" 
 script which pulls data from the appropriate Excel XML file. With this data, 
 I added to the script to run a macro on the stepper tool to raw_input this data.
"""

import os
import pandas as pd
import time
import math
from pynput.keyboard import Key, Controller


# User Parameters/Constants to Set
XML_DIR = "./XML_Files/"
SLEEP_TIME_BETWEEN = 0.01
SLEEP_TIME_BETWEEN_LETTER = 0.01
SLEEP_TIME_END = 1 


def time_convert(sec):
  mins = sec // 60
  sec = sec % 60
  hours = mins // 60
  mins = mins % 60
  print("Time Lapsed = {0}h:{1}m:{2}s".format(int(hours) ,int(mins), round(sec)))


# Presses Alt + Tab
def altTab():
    time.sleep(SLEEP_TIME_BETWEEN)
    with keyboard.pressed(Key.alt):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)


# Type's in string
def typeThis(toType):
    time.sleep(SLEEP_TIME_BETWEEN)
    for letter in toType:
        keyboard.type(letter)
        time.sleep(SLEEP_TIME_BETWEEN_LETTER)
    time.sleep(SLEEP_TIME_BETWEEN)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(SLEEP_TIME_BETWEEN)



# MAIN():
# =============================================================================

# Starting stopwatch to see how long process takes
start_time = time.time()

# Clears some of the screen for asthetics
print("\n\n\n\n\n\n\n\n\n\n\n\n\n")

keyboard = Controller()

# Runs through each XML file in folder placement
for XML_Name in os.listdir(XML_DIR):
    
    # Skips non Excel files
    if ".xlsx" not in XML_Name:
        continue
    
    XML_Path = os.path.join(XML_DIR, XML_Name)
    
    # Pulls data from XML file
    # -------------------------------------------------------------------------
    # Reads from first tab, "Alignment and Main" and grabs values
    df = pd.read_excel(XML_Path, sheet_name='Alignment and Main')
    
    step_x      = df['Value'][0]
    stepdist_x  = df['Value'][1]
    step_y      = df['Value'][2]
    stepdist_y  = df['Value'][3]
    lkey_R      = df['Value'][4]
    lkey_C      = df['Value'][5]
    lkey_x      = df['Value'][6]
    lkey_y      = df['Value'][7]
    rkey_R      = df['Value'][8]
    rkey_C      = df['Value'][9]
    rkey_x      = df['Value'][10]
    rkey_y      = df['Value'][11]
    # Below changed for new excel sheet
    # mask_x      = df['Value'][12]
    # mask_y      = df['Value'][13]
    # wafer_x     = df['Value'][14]
    # wafer_y     = df['Value'][15]

    wafer_size  = df['Value'][12] # in mm
    # left_blade  = df['Value'][19]
    # right_blade = df['Value'][20]
    # rear_blade  = df['Value'][21]
    # front_blade = df['Value'][22]
    
    
    rkey_xoffset = round(((-1*(step_x/2-0.5)*stepdist_x+(step_x-rkey_C)*stepdist_x) - rkey_x) / 1000, 5)
    rkey_yoffset = round(((rkey_y - (-1*(step_y/2-0.5)*stepdist_y+(rkey_R-1)*stepdist_y) ) / 1000), 5)
    lkey_xoffset = round((((-1*(step_x/2-0.5)*stepdist_x+(step_x-lkey_C)*stepdist_x) - lkey_x) / 1000), 5)
    lkey_yoffset = round(((lkey_y - (-1*(step_y/2-0.5)*stepdist_y+(lkey_R-1)*stepdist_y)) / 1000), 5)
    
    # pshift_x = round((((-1*(step_x/2-0.5)*stepdist_x) + mask_x) - wafer_x) / 1000, 8)
    # pshift_y = round((wafer_y - ((-1*(step_y/2-0.5)*stepdist_y) + mask_y) ) / 1000, 8)

    
    # left_b = (left_blade * 5) / 1000 + 50
    # right_b = 50 - (right_blade * 5) / 1000
    # front_b = (front_blade * 5) / 1000 + 50
    # rear_b = 50 - (rear_blade * 5) / 1000
    
    
    # REEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
    altTab()
    # UPDATE CREATION DATE: 
    typeThis("Y")
    # Job Comment: 
    typeThis("")
    # Tolerance:
    typeThis("3")
    # Scale Corrections:
    typeThis("0")
    typeThis("0")
    # Orthoganality:
    typeThis("0")
    # Leveler Batch Size:
    typeThis("1")
    
    # Wafer Diameter:
    if wafer_size == 200:
        typeThis("215")
    elif wafer_size == 150:
        typeThis("180")
    elif wafer_size == 100:
        typeThis("160")
    else:
        break
    
    # Step Size: X:
    typeThis(str(round(stepdist_x / 1000, 5)))
    # C-ount, S-Pan, A-All:
    typeThis("C")
    # How Many Columns?
    typeThis(str(int(step_x)))
    # Step Size: Y:
    typeThis(str(round(stepdist_y / 1000, 5)))
    # C-ount, S-Pan, A-All:
    typeThis("C")
    # How Many Rows?:
    typeThis(str(int(step_y)))
    
    # Translate Origin
    if wafer_size == 100:
        # X: 
        typeThis("-18")
        # Y: 
        typeThis("20")
    else:
        # X: 
        typeThis("0")
        # Y: 
        typeThis("0")
    
    # Display?: 
    typeThis("N")
    # Layout?: 
    typeThis("N")
    # Adjust?: 
    typeThis("N")
    
    # Alignment Parameters:
    # Standard Keys?:
    typeThis("N")
    # Right Alignment Die Center:
    # R:
    typeThis(str(int(rkey_R)))
    # C: 
    typeThis(str(int(rkey_C)))
    # Right Key Offset:
    # X:
    typeThis(str(round(rkey_xoffset,5)))
    # Y:
    typeThis(str(round(rkey_yoffset,5)))
    # Left Alignment Die Center: 
    # R: 
    typeThis(str(int(lkey_R)))
    # C: 
    typeThis(str(int(lkey_C)))
    # Left Key Offset: 
    # X: 
    typeThis(str(round(lkey_xoffset,5)))
    # Y: 
    typeThis(str(round(lkey_yoffset,5)))
    # EPI Shift:
    # X: 
    typeThis("")
    # Y: 
    typeThis("")
    
    # TO NEW PASS OR ARRAY
    # Name:
    typeThis("")
    # Main Array: 
    # Pass Comment: 
    # Main Array
    typeThis("")
    # Use Local Alignment?:
    typeThis("N")
    # Exposure (Sec): 
    typeThis("")
    # Exposure Scale Factor [0 -> 2]: 
    typeThis("")
    # Focus Offset: 
    typeThis("")
    # Microsoft Focus Offset [-2000 -> +2000]: 
    typeThis("")
    # Enable Match?:
    typeThis("N")
    # Pass Shift: 
    # X: 
    typeThis(str(pshift_x))
    # Y: 
    typeThis(str(pshift_y))
    # Recticle Bar Code: 
    typeThis("NONE")

    # Masking Arperture Settings: 
    # XL: 
    typeThis(str(round(left_b, 5)))
    # XR: 
    typeThis(str(round(right_b, 5)))
    # YF: 
    typeThis(str(round(front_b, 5)))
    # YR: 
    typeThis(str(round(rear_b, 5)))
    
    # Recticle Alignment Offset: 
    # XL: 
    typeThis("0")
    # XR: 
    typeThis("0")
    # Y: 
    typeThis("0")
    
    # Reticle Alignment Mark Phase (P, N, X): 
    typeThis("N")
    # Reticle Transmission % [ 0 - 300 ]:
    typeThis("")
    
    
    
    break
    # REEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
    
    
    i = 0
    while True:
        df_p_i = pd.read_excel(XML_Path, sheet_name = i+2) 
        if math.isnan(df_p_i['Number of Plugs'][0]):
            break
        #print('\n', 'Pass Name:', df_p_i['Pass Name'][0], '\n')
        typePass = df_p_i['Plugs or Array'][0]
        test_mask_x = df_p_i['Test site bottom left x on mask'][0]
        test_mask_y = df_p_i['Test site bottom left y on mask'][0]
        left_blade  = df_p_i['Plug Pass Reticle Blade Left Position'][0]
        right_blade = df_p_i['Plug Pass Reticle Blade Right Position'][0]
        front_blade = df_p_i['Plug Pass Reticle Blade Bottom Position'][0]
        rear_blade  = df_p_i['Plug Pass Reticle Blade Top Position'][0]
        
        left_b  = (left_blade * 5) / 1000 + 50
        right_b = 50 - (right_blade * 5) / 1000
        front_b = (front_blade * 5) / 1000 + 50
        rear_b  = 50 - (rear_blade * 5) / 1000
        
        if typePass == "A":
            wafer_x = df_p_i['Bottom Left x on wafer'][0]
            wafer_y = df_p_i['Bottom Left y on wafer'][0]
            pshift_x = round((((-1*(step_x/2-0.5)*stepdist_x) + test_mask_x) - wafer_x) / 1000, 8)
            pshift_y = round((wafer_y - ((-1*(step_y/2-0.5)*stepdist_y) + test_mask_y) ) / 1000, 8)
            print('\nPASS SHIFT:\nX:', pshift_x, '\nY:', pshift_y, '\n')  
            
            # Name:
            typeThis("")
            # Main Array: 
            # Pass Comment: 
            # Main Array
            typeThis("")
            # Use Local Alignment?:
            typeThis("N")
            # Exposure (Sec): 
            typeThis("")
            # Exposure Scale Factor [0 -> 2]: 
            typeThis("")
            # Focus Offset: 
            typeThis("")
            # Microsoft Focus Offset [-2000 -> +2000]: 
            typeThis("")
            # Enable Match?:
            typeThis("N")
            # Pass Shift: 
            # X: 
            typeThis(str(pshift_x))
            # Y: 
            typeThis(str(pshift_y))
            # Recticle Bar Code: 
            typeThis("NONE")
        
            # Masking Arperture Settings: 
            # XL: 
            typeThis(str(round(left_b, 5)))
            # XR: 
            typeThis(str(round(right_b, 5)))
            # YF: 
            typeThis(str(round(front_b, 5)))
            # YR: 
            typeThis(str(round(rear_b, 5)))
            
            # Recticle Alignment Offset: 
            # XL: 
            typeThis("0")
            # XR: 
            typeThis("0")
            # Y: 
            typeThis("0")
            
            # Reticle Alignment Mark Phase (P, N, X): 
            typeThis("N")
            # Reticle Transmission % [ 0 - 300 ]:
            typeThis("")
        
        elif typePass == "P":
            # Name:
            typeThis("")
            # Main Array: 
            # Pass Comment: 
            # Main Array
            typeThis("")
            # Use Local Alignment?:
            typeThis("N")
            # Exposure (Sec): 
            typeThis("")
            # Exposure Scale Factor [0 -> 2]: 
            typeThis("")
            # Focus Offset: 
            typeThis("")
            # Microsoft Focus Offset [-2000 -> +2000]: 
            typeThis("")
            # Enable Match?:
            typeThis("N")
            # Pass Shift: 
            # X: 
            typeThis("")
            # Y: 
            typeThis("")
            # Recticle Bar Code: 
            typeThis("NONE")
        
            # Masking Arperture Settings: 
            # XL: 
            typeThis(str(round(left_b, 5)))
            # XR: 
            typeThis(str(round(right_b, 5)))
            # YF: 
            typeThis(str(round(front_b, 5)))
            # YR: 
            typeThis(str(round(rear_b, 5)))
            
            # Recticle Alignment Offset: 
            # XL: 
            typeThis("0")
            # XR: 
            typeThis("0")
            # Y: 
            typeThis("0")
            
            # Reticle Alignment Mark Phase (P, N, X): 
            typeThis("N")
            # Reticle Transmission % [ 0 - 300 ]:
            typeThis("")


    
        #print('Aperture Blades:', '\n', 'Left:', left_b, '\n', 'Right:', right_b, '\n', 'Front:', front_b, '\n', 'Rear:', rear_b,)
        num_plugs = int(df_p_i['Number of Plugs'][0])
        
        if typePass == "A":
            typeThis("A")
        elif typePass == "P":
            typeThis("P")
        
        if typePass == "P":
            for x in range(num_plugs):
                x_offset = round((((-1*(step_x/2-0.5)*stepdist_x+(step_x-df_p_i['Plug Closest Column'][x])*stepdist_x) + test_mask_x) - 
                df_p_i['Bottom Left x on wafer'][x]) / 1000, 5)
                y_offset = round((df_p_i['Bottom Left y on wafer'][x] - ((-1*(step_y/2-0.5)*stepdist_y+
                                                                   (df_p_i['Plug Closest Row '][x]-1)*stepdist_y)
                                                                  + test_mask_y)) / 1000, 5)
                # print('\n', 'Plug', str(df_p_i['Plug Number '][x]) + ':', '\n', 'R:', df_p_i['Plug Closest Row '][x], '\t', 'y:', 
                #   y_offset, '\n',
                #   'C:', df_p_i['Plug Closest Column'][x], '\t', 'x:', x_offset)
                
                # REEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
                typeThis(str(df_p_i['Plug Number '][x]))
                typeThis(str(df_p_i['Plug Closest Row '][x]))
                typeThis(str(y_offset))
                typeThis(str(df_p_i['Plug Closest Column'][x]))
                typeThis(str(x_offset))
                # REEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
        elif typePass == "A":
            # SOMETHING BAOUT DROPOUTS
            typeThis("")
            print("yayyyyy")
            
            
        
        i += 1


            
    #print('PASS\nNAME: MAP\nPASS COMMENT:\nMAPPING')
    #print('USE LOCAL ALIGNMENT?: Y')
    #print('EXPOSE MAPPING PASS?: N\nUSE TWO POINT ALIGNMENT?:\nNUMBER OF ALIGNMENTS PER DIE?:1')
    #print('LOCAL ALIGNMENT MARK OFFSET:\nX:0\nY:0\nMONITOR MAPPING CORRECTIONS?:Y')
    #print('MAP EVERY N TH WAFER N = 1\nMICROSCOPE FOCUS OFFSET:0\nPASS SHIFT:\nX:0\nY:0\nA-RRAY OR P-LUG: P')

    #print('\n\nPLUGS:\n')
    
    df_m = pd.read_excel(XML_Path, sheet_name = 'Mapping')
    i = 0
    
    # REEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
    test = raw_input("Press enter to continue auto typing second set.")
    altTab()
    # REEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
    
    while True:
        if math.isnan(df_m['Closest Column'][i]):
            break
        x_offset = round((((-1*(step_x/2-0.5)*stepdist_x+(step_x-df_m['Closest Column'][i])*stepdist_x) - 
                          df_m['Center of alignment site x on wafer'][i]) / 1000), 5)
        y_offset = round(((df_m['Center of alignment site y on wafer'][i] -
                           (-1*(step_y/2-0.5)*stepdist_y+(df_m['Closest Row'][i]-1)*stepdist_y) 
                           )) / 1000, 5)
        #print('Plug', str(df_m['Map Site Number'][i]) + ':', '\n', 'R:', df_m['Closest Row'][i], '\t', 'y:', y_offset, '\n', 
             # 'C:', df_m['Closest Column'][i], '\t', 'x:', x_offset)
        
        # REEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
        
        typeThis(str(df_m['Map Site Number'][i]))
        typeThis(str(df_m['Closest Row'][i]))
        typeThis(str(y_offset))
        typeThis(str(df_m['Closest Column'][i]))
        typeThis(str(x_offset))
        # REEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
        
        i += 1
        
        
      
    # REEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
    altTab()
    # REEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
    
    #print('NAME (<CR> TO EXIT PASS SETUP) :')
    #print('WRITE TO DISK?:Y')
    #print('PURGE EDITED FILES ? : Y')
    # -------------------------------------------------------------------------
    
    
# Macro
# -----------------------------------------------------------------------------






        

# -----------------------------------------------------------------------------

print("\nThis program is done!")

# Starting stopwatch to see how long process takes
print("Total Time: ")
end_time = time.time()
time_lapsed = end_time - start_time
time_convert(time_lapsed)
# =============================================================================