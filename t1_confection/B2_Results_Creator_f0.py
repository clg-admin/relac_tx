# -*- coding: utf-8 -*-
# Suggested Citation:
# Victor-Gallardo, L. (2022). Robust Energy System Planning for Decarbonization under Technological Uncertainty.
# Retrieved from https://www.kerwa.ucr.ac.cr/handle/10669/87273

"""
This code is based on the methodologies and findings discussed in the work of Luis Victor-Gallardo, specifically focusing on robust energy system planning for decarbonization under technological uncertainty. The document provides a comprehensive analysis relevant to the code's development and application. For in-depth understanding and further details, please refer to the original document linked below.

Suggested Citation:
Victor-Gallardo, L. (2022). Robust Energy System Planning for Decarbonization under Technological Uncertainty.
University of Costa Rica. Retrieved from https://www.kerwa.ucr.ac.cr/handle/10669/87273
"""

from datetime import date
import sys
import pandas as pd
import os
from copy import deepcopy
import csv
import numpy as np
import yaml
import platform

sys.path.insert(0, 'Executables')
import local_dataset_creator_0

def get_config_main_path(full_path, base_folder='config_main_files'):
    # Split the path into parts
    parts = full_path.split(os.sep)
    
    # Find the index of the target directory 'relac_tx'
    target_index = parts.index('relac_tx') if 'relac_tx' in parts else None
    
    # If the directory is found, reconstruct the path up to that point
    if target_index is not None:
        base_path = os.sep.join(parts[:target_index + 1])
    else:
        base_path = full_path  # If not found, return the original path
    
    # Append the specified directory to the base path
    appended_path = os.path.join(base_path, base_folder)
    
    return appended_path

def load_and_process_yaml(path):
    """
    Load a YAML file and replace the specific placeholder '${year_apply_discount_rate}' with the year specified in the file.
    
    Args:
    path (str): The path to the YAML file.
    
    Returns:
    dict: The updated data from the YAML file where the specific placeholder is replaced.
    """
    with open(path, 'r') as file:
        # Load the YAML content into 'params'
        params = yaml.safe_load(file)

    # Retrieve the reference year from the YAML file and convert it to string for replacement
    reference_year = str(params['year_apply_discount_rate'])

    # Function to recursively update strings containing the placeholder
    def update_strings(obj):
        if isinstance(obj, dict):
            return {k: update_strings(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [update_strings(element) for element in obj]
        elif isinstance(obj, str):
            # Replace the specific placeholder with the reference year
            return obj.replace('${year_apply_discount_rate}', reference_year)
        else:
            return obj

    # Update all string values in the loaded YAML structure
    updated_params = update_strings(params)

    return updated_params

# Read yaml file with parameterization
file_config_address = get_config_main_path(os.path.abspath(''))
params = load_and_process_yaml(os.path.join(file_config_address, 'MOMF_B1_exp_manager.yaml'))

#------------------------------------------------------------------------------------------------
# For macOS system delete a folder hidden
if platform.system() == 'Darwin':
    os.remove(os.path.join('Executables', '.DS_Store'))
#------------------------------------------------------------------------------------------------

# Assuming `df` is your DataFrame after loading the Excel file
df_structure = pd.read_excel('B1_Model_Structure.xlsx')

# Now let's find the last non-empty value in the 'REGION' column
region_str = df_structure['REGION'].dropna().iloc[-1]

file_name_suffix = 'f0_OSMOSYS_' + region_str + '_'

run_for_first_time = True

if run_for_first_time == True:
    local_dataset_creator_0.execute_local_dataset_creator_0_outputs()
    local_dataset_creator_0.execute_local_dataset_creator_0_inputs()
############################################################################################################
path_output_0 = os.path.join('Executables', 'output_dataset_0.csv')
df_0_output = pd.read_csv(path_output_0, index_col=None, header=0, low_memory=False)
#
li_output = [df_0_output]
#
df_output = pd.concat(li_output, axis=0, ignore_index=True)
if params['solver']=='glpk' and params['glpk_option']=='old':
    df_output.sort_values(by=['Future.ID','Fuel','Technology','Emission','Year'], inplace=True)
else:
    df_output.sort_values(by=['FUEL','TECHNOLOGY','EMISSION','YEAR'], inplace=True)


sector_flag = False
if sector_flag:
    df_output=df_output.assign(Sector=np.nan)
    df_output=df_output.assign(Description=np.nan)
    df_output=df_output.assign(SpecificSector=np.nan)


    libro = pd.ExcelFile('B1_Model_Structure.xlsx')
    hoja=libro.parse( 'sector' , skiprows = 0 )
    encabezados=list(hoja)

    col_t=list(hoja[encabezados[0]])
    col_s=list(hoja[encabezados[1]])
    col_d=list(hoja[encabezados[2]])
    col_ss=list(hoja[encabezados[3]])
    dicSector=dict(zip(col_t,col_s))
    dicDescription=dict(zip(col_t,col_d))
    dicSpecificSector=dict(zip(col_t,col_ss))

    llaves=list(dicSector.keys())
    col_sector=np.array(list(df_output['Sector']))

    for i in range(len(llaves)):
        df_output.loc[df_output['TECHNOLOGY'] == llaves[i], 'Sector'] =  dicSector[llaves[i]]
        df_output.loc[df_output['TECHNOLOGY'] == llaves[i], 'Description'] =  dicDescription[llaves[i]]
        df_output.loc[df_output['TECHNOLOGY'] == llaves[i], 'SpecificSector'] =  dicSpecificSector[llaves[i]]
        

############################################################################################################
path_input_0 = os.path.join('Executables', 'input_dataset_0.csv')
df_0_input = pd.read_csv(path_input_0, index_col=None, header=0, low_memory=False)
#
li_intput = [df_0_input]
#
df_input = pd.concat(li_intput, axis=0, ignore_index=True)
df_input.sort_values(by=['Future.ID','Strategy.ID','Strategy','Fuel','Technology','Emission','Season','Year'], inplace=True)

############################################################################################################
#
dfa_list = [df_output, df_input]
#
today = date.today()
#
df_output = dfa_list[0]
df_output.to_csv (region_str + 'Output.csv', index = None, header=True)
df_output.to_csv (region_str + 'Output_' + str( today ).replace( '-', '_' ) + '.csv', index = None, header=True)
#
df_input = dfa_list[1]
df_input.to_csv (region_str + 'Input.csv', index = None, header=True)
df_input.to_csv (region_str + 'Input_' + str( today ).replace( '-', '_' ) + '.csv', index = None, header=True)
#