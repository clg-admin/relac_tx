# -*- coding: utf-8 -*-
"""
Created on Thu Feb 22 10:52:32 2024

@author: andreysava19
"""

import yaml
import os
from copy import deepcopy
import shutil
import sys

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


if __name__ == '__main__':
    # Read yaml file with parameterization
    file_config_address = get_config_main_path(os.path.abspath(''))
    params = load_and_process_yaml(os.path.join(file_config_address, 'MOMF_B1_exp_manager.yaml'))
    # Select dict with default values parameters
    default_val_params = params['default_val_params']
    default_val_sets = params['sets_otoole']
    # Parameters otoole has by default but don't use them in MOMF
    parameters_otoole_no_momf = params['parameters_otoole_no_momf']
    
    # Create defaul yaml file by otoole
    script_conv_format = os.path.join('config', params['conv_format'])
    file_path = os.path.join(file_config_address, script_conv_format)
    directory = os.path.dirname(file_path)
    if not os.path.exists(directory):
        os.makedirs(directory)
        
    os.system(f'otoole setup --overwrite config {file_path}')
    
    # Create folder with csv templates
    script_templates = os.path.join('config', params['templates'])
    file_path_templates = os.path.join(file_config_address, script_templates)
    os.system(f'otoole setup --overwrite csv {file_path_templates}')
    
    # Read yaml file with parameterization to conversion format by otoole
    with open(file_path, 'r') as file:
        # Load content file
        default_format = yaml.safe_load(file)
    
    # Create temporal variable to delete some parameter from 'default_format'
    temp_default_params = deepcopy(default_format)
    
    # Delete parameters that are not in 'MOMF_T1_B1.yaml'
    for old_key, attributes in temp_default_params.items():
        script_templates_old_key = os.path.join(script_templates, old_key)
        script_templates_old_key = script_templates_old_key + '.csv'
        file_to_delete = os.path.join(file_config_address, script_templates_old_key)
        if attributes['type'] == 'param':
            for index in attributes['indices']:
                if index not in default_val_sets:
                    default_format[old_key]['indices'].remove(index)
                        
            if not old_key in default_val_params:
                default_format.pop(old_key)
                shutil.os.remove(file_to_delete)
            
        elif attributes['type'] == 'set' and not old_key in default_val_sets:
            a = default_format.pop(old_key)
            shutil.os.remove(file_to_delete)
        
        elif attributes['type'] == 'result':
            for index in attributes['indices']:
                if index not in default_val_sets:
                    default_format[old_key]['indices'].remove(index)
                    # shutil.os.remove(file_config_address + 'config\\templates\\' + index + '.csv')      
    
    # Delete key asociate with the variable "parameters_otoole_no_momf"
    cleaned_default_format = {k: v for k, v in default_format.items() if k not in parameters_otoole_no_momf}
    
    # Add new parameters need for MOMF
    for key, value in params['new_parameters_add'].items():
        if key not in cleaned_default_format:
            cleaned_default_format[key] = value
    
    # Assing dict clean of the parameters aren't useful to default_format dict
    default_format = cleaned_default_format
        
                
    # Change default value
    for old_key, attributes in default_format.items():
        if attributes['type'] == 'param' and 'default' in attributes:
            default_val_params = params['default_val_params']
            default_format[old_key]['default'] = default_val_params[old_key]
            
    # Save changes of default values in the yaml file
    with open(file_path, 'w') as file:
        yaml.dump(default_format, file, sort_keys=False)