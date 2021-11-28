"""
This python script uses an edited version of Brady Pierce's "Stepper_Job.py" 
 script which pulls data from the appropriate Excel XML file. With this data, 
 I added to the script to run a macro on the stepper tool to input this data.
"""

import os
import pandas as pd
import time
import math
import sys
from subprocess import Popen, PIPE, STDOUT


# User Parameters/Constants to Set
XML_DIR = "./XML_Files/"
NUM_MAP_SITES = 6


def time_convert(sec):
  mins = sec // 60
  sec = sec % 60
  hours = mins // 60
  mins = mins % 60
  print("Time Lapsed = {0}h:{1}m:{2}s".format(int(hours) ,int(mins), round(sec)))



# MAIN():
# =============================================================================

# Starting stopwatch to see how long process takes
start_time = time.time()

# Clears some of the screen for asthetics
print("\n\n\n\n\n\n\n\n\n\n\n\n\n")

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
    mask_x      = df['Value'][12]
    mask_y      = df['Value'][13]
    wafer_x     = df['Value'][14]
    wafer_y     = df['Value'][15]

    wafer_size  = df['Value'][18] # in mm
    left_blade  = df['Value'][19]
    right_blade = df['Value'][20]
    rear_blade  = df['Value'][21]
    front_blade = df['Value'][22]
    
    print('UPDATE CREATION DATE: Y')
    print('JOB COMMENT: ' + 'job name')
    print('TOLERANCE: 3')
    print('SCALE CORRECTIONS \n X, PPM: 0 \n Y, PPM: 0 \n ORTHOGONALITY, PPM:0 \n LEVELER BATCH SIZE:1')
    
    
    if wafer_size == 200:
        print('Wafer diameter: 215')
    elif wafer_size == 150:
        print('Wafer diameter: 180')
    elif wafer_size == 100:
        print('\n', 'Wafer diameter: 160')
    else:
        print('Unsupported wafer size')
    
    print('\n', 'STEP SIZE:', '\n', 'X:', stepdist_x / 1000, '\n COUNT', '\n HOW MANY COLUMNS?', step_x, '\n', 'STEP SIZE \n Y:', 
          stepdist_y / 1000, '\n COUNT', '\n HOW MANY ROWS?', step_y, '\n')
    
    if wafer_size == 100:
        print('TRANSLATE ORIGIN:', '\n', 'X: -18', '\n', 'Y: 20', '\n')
    else:
        print('TRANSLATE ORIGIN:', '\n', 'X: 0', '\n', 'Y: 0', '\n')
        
    print('DISPLAY?: N \nLAYOUT?: N \nADJUST: N')

    rkey_xoffset = round(((-1*(step_x/2-0.5)*stepdist_x+(step_x-rkey_C)*stepdist_x) - rkey_x) / 1000, 5)
    rkey_yoffset = round(((rkey_y - (-1*(step_y/2-0.5)*stepdist_y+(rkey_R-1)*stepdist_y) ) / 1000), 5)
    lkey_xoffset = round((((-1*(step_x/2-0.5)*stepdist_x+(step_x-lkey_C)*stepdist_x) - lkey_x) / 1000), 5)
    lkey_yoffset = round(((lkey_y - (-1*(step_y/2-0.5)*stepdist_y+(lkey_R-1)*stepdist_y)) / 1000), 5)
    print ('ALIGNMENT PARAMATERS \nSTANDARD KEYS?: N\nRIGHT ALIGNMENT DIE CENTER:', '\n', 'R:', int(rkey_R), '\n', 'C:', int(rkey_C)
           , '\nRIGHT KEY OFFSET\nX:', rkey_xoffset, '\nY:', rkey_yoffset, '\nLEFT ALGINMENT DIE CENTER:', '\n', 'R:', int(lkey_R), '\n', 
           'C:', int(lkey_C), '\nLEFT KEY OFFSET\nX:', lkey_xoffset, '\nY:', lkey_yoffset, '\n')
    
    print('EPI SHIFT\nX:\nY:')
    

    print('\n\nPASS')
    #FIX#
    print('\nNAME:', 'name of pass')
    print('\nPASS COMMENT:\n', 'name of pass comment')
    print('\nUSE LOCAL ALIGNMENT?: N')
    print('ENABLE MATCH ?: N')
    
    pshift_x = round((((-1*(step_x/2-0.5)*stepdist_x) + mask_x) - wafer_x) / 1000, 8)
    pshift_y = round((wafer_y - ((-1*(step_y/2-0.5)*stepdist_y) + mask_y) ) / 1000, 8)
    print('\nPASS SHIFT:\nX:', pshift_x, '\nY:', pshift_y, '\n')

    
    left_b = (left_blade * 5) / 1000 + 50
    right_b = 50 - (right_blade * 5) / 1000
    front_b = (front_blade * 5) / 1000 + 50
    rear_b = 50 - (rear_blade * 5) / 1000
    print('MASKING APERTURE SETTINGS:\nXL:', left_b, '\nXR:', right_b, '\nYF:', front_b, '\nYR:', rear_b)
    print('RETICLE ALIGNMENT OFFSET (MICRONS) :\nXL:0\nXR:0\nY:0')
    print('RETICLE ALIGNMENT MARK PHASE: N')
    print('A-RRAY OR P-LUG:', 'array')
    
    print('DROPOUTS:', 'see dropouts')
    
    print('SAVE PASS? : Y')
    
    i = 0
    while True:
        df_p_i = pd.read_excel(XML_Path, sheet_name = i+2) 
        if math.isnan(df_p_i['Number of Plugs'][0]):
            break
        print('\n', 'Pass Name:', df_p_i['Pass Name'][0], '\n')
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
    
        print('Aperture Blades:', '\n', 'Left:', left_b, '\n', 'Right:', right_b, '\n', 'Front:', front_b, '\n', 'Rear:', rear_b,)
        num_plugs = int(df_p_i['Number of Plugs'][0])
        
        for x in range(num_plugs):
            x_offset = round((((-1*(step_x/2-0.5)*stepdist_x+(step_x-df_p_i['Plug Closest Column'][x])*stepdist_x) + test_mask_x) - 
            df_p_i['Bottom Left x on wafer'][x]) / 1000, 5)
            y_offset = round((df_p_i['Bottom Left y on wafer'][x] - ((-1*(step_y/2-0.5)*stepdist_y+
                                                               (df_p_i['Plug Closest Row '][x]-1)*stepdist_y)
                                                              + test_mask_y)) / 1000, 5)
            print('\n', 'Plug', str(df_p_i['Plug Number '][x]) + ':', '\n', 'R:', df_p_i['Plug Closest Row '][x], '\t', 'y:', 
              y_offset, '\n',
              'C:', df_p_i['Plug Closest Column'][x], '\t', 'x:', x_offset)
        
        i += 1
            
            
    print('PASS\nNAME: MAP\nPASS COMMENT:\nMAPPING')
    print('USE LOCAL ALIGNMENT?: Y')
    print('EXPOSE MAPPING PASS?: N\nUSE TWO POINT ALIGNMENT?:\nNUMBER OF ALIGNMENTS PER DIE?:1')
    print('LOCAL ALIGNMENT MARK OFFSET:\nX:0\nY:0\nMONITOR MAPPING CORRECTIONS?:Y')
    print('MAP EVERY N TH WAFER N = 1\nMICROSCOPE FOCUS OFFSET:0\nPASS SHIFT:\nX:0\nY:0\nA-RRAY OR P-LUG: P')

    print('\n\nPLUGS:\n')
    
    df_m = pd.read_excel(XML_Path, sheet_name = 'Mapping')
    i = 0
    while True:
        if math.isnan(df_m['Closest Column'][i]):
            break
        x_offset = round((((-1*(step_x/2-0.5)*stepdist_x+(step_x-df_m['Closest Column'][i])*stepdist_x) - 
                          df_m['Center of alignment site x on wafer'][i]) / 1000), 5)
        y_offset = round(((df_m['Center of alignment site y on wafer'][i] -
                           (-1*(step_y/2-0.5)*stepdist_y+(df_m['Closest Row'][i]-1)*stepdist_y) 
                           )) / 1000, 5)
        print('Plug', str(df_m['Map Site Number'][i]) + ':', '\n', 'R:', df_m['Closest Row'][i], '\t', 'y:', y_offset, '\n', 
             'C:', df_m['Closest Column'][i], '\t', 'x:', x_offset)
        i += 1
        
        
    print('NAME (<CR> TO EXIT PASS SETUP) :')
    print('WRITE TO DISK?:Y')
    print('PURGE EDITED FILES ? : Y')
    # -------------------------------------------------------------------------
    
    
    # Macro
    # -------------------------------------------------------------------------



print("\nThis program is done!")

# Starting stopwatch to see how long process takes
print("Total Time: ")
end_time = time.time()
time_lapsed = end_time - start_time
time_convert(time_lapsed)
# =============================================================================