# -*- coding: utf-8 -*-
"""
Created on Mon May 26 17:22:31 2025

@author: ClimateLeadGroup
"""

import os
import pandas as pd
import yaml
import subprocess
import sys
import platform  
import shutil
import time

########################################################################################
def sort_csv_files_in_folder(folder_path):
    if not os.path.isdir(folder_path):
        print(f"The path is invalid: {folder_path}")
        return
    print('################################################################')
    print('Sort csv files.')
    for filename in os.listdir(folder_path):
        if filename.endswith(".csv"):
            file_path = os.path.join(folder_path, filename)
            print(f"Processing: {filename}")
            try:
                # Leer el CSV preservando la cabecera
                df = pd.read_csv(file_path)

                # Ordenar usando todas las columnas
                df_sorted = df.sort_values(by=list(df.columns))

                # Sobrescribir el archivo original
                df_sorted.to_csv(file_path, index=False)
            except Exception as e:
                print(f"Error processing {filename}: {e}")

    print("✅ All files were sort.")
    print('################################################################\n')

def process_scenario_folder(base_input_path, template_path, base_output_path, scenario_name):
    """
    Processes a scenario folder: reads its CSV files, aligns with template structure,
    maps 'Value' to 'VALUE', excludes specific columns, and saves the results to output.
    Also ensures VALUE is int() for certain template files.
    """

    # Step 1: Define scenario input path
    scenario_input_path = os.path.join(base_input_path, scenario_name)

    # Step 2: Skip if not a directory or is 'Default'
    if not os.path.isdir(scenario_input_path) or scenario_name == 'Default':
        return

    # Step 3: Read and clean scenario CSVs
    scenario_files = {}
    for f in os.listdir(scenario_input_path):
        if f.endswith('.csv'):
            df = pd.read_csv(os.path.join(scenario_input_path, f))

            # Remove unwanted columns
            df = df.drop(columns=[col for col in ['PARAMETERT', 'Scenario'] if col in df.columns])
            df = df.dropna(axis=1, how='all')

            # Rename 'Value' to 'VALUE'
            if 'Value' in df.columns:
                df = df.rename(columns={'Value': 'VALUE'})

            scenario_files[f] = df

    # Step 4: Read template files
    template_files = {
        f: pd.read_csv(os.path.join(template_path, f))
        for f in os.listdir(template_path)
        if f.endswith('.csv')
    }

    # Step 5: Create output path
    scenario_output_path = os.path.join(base_output_path, scenario_name)
    os.makedirs(scenario_output_path, exist_ok=True)
    
    # Step 6: Fill templates with scenario data
    for template_name, template_df in template_files.items():
        output_file_path = os.path.join(scenario_output_path, template_name)
        
        if template_name in scenario_files:
            input_df = scenario_files[template_name]
            common_columns = [col for col in template_df.columns if col in input_df.columns]
            filled_df = template_df.copy()
            filled_df[common_columns] = input_df[common_columns]

            # Step 7: Convert VALUE to int if required
            if template_name in [
                'DAYTYPE.csv', 'DAILYTIMEBRACKET.csv', 'SEASON.csv',
                'MODE_OF_OPERATION.csv', 'YEAR.csv', 'EMISSION.csv',
                'FUEL.csv', 'REGION.csv', 'STORAGE.csv', 'TECHNOLOGY.csv',
                'TIMESLICE.csv', 'Conversionls.csv'
            ]:
                if 'VALUE' in filled_df.columns:
                    # Drop rows with NaN or empty string (including whitespace-only)
                    filled_df = filled_df[filled_df['VALUE'].notna() & (filled_df['VALUE'].astype(str).str.strip() != '')]
            
                    # Convert to int if required
                    if template_name in [
                        'DAYTYPE.csv', 'DAILYTIMEBRACKET.csv', 'SEASON.csv',
                        'MODE_OF_OPERATION.csv', 'YEAR.csv'
                    ]:
                        filled_df['VALUE'] = filled_df['VALUE'].astype(int)

            filled_df.to_csv(output_file_path, index=False)
        else:
            template_df.to_csv(output_file_path, index=False)
            
    folder_to_sort = os.path.join(base_output_path,scenario_name)
    sort_csv_files_in_folder(folder_to_sort)

    print(f"✅ Scenario '{scenario_name}': templates filled and saved successfully.\n")
    print('#------------------------------------------------------------------------------#')

def run_otoole_conversion(base_output_path, scenario_name, params):
    """
    Runs the corrected 'otoole convert csv datafile' command for a given scenario.

    Parameters:
        base_output_path (str): Path where the scenario's filled CSV files are stored.
        scenario_name (str): The name of the scenario.
        params (dict): Dictionary loaded from the YAML file with required paths.
    """
    # Step 1: Define paths
    input_folder = os.path.join(base_output_path, scenario_name)
    scenario_exec_dir = os.path.join(params['executables'], scenario_name + '_0')
    output_file = os.path.join(scenario_exec_dir, f"{scenario_name}_0.txt")
    config_file = os.path.join(params['Miscellaneous'], params['otoole_config'])

    # Step 2: Ensure the scenario's executable folder exists
    os.makedirs(scenario_exec_dir, exist_ok=True)

    # Step 3: Construct the command
    command = [
        'otoole', 'convert', 'csv', 'datafile',
        input_folder,
        output_file,
        config_file
    ]

    print(f"Running command: {' '.join(command)}")

    # Step 4: Execute the command
    result = subprocess.run(command, capture_output=True, text=True)

    # Step 5: Handle output
    if result.returncode != 0:
        print(f"❌ Error while converting scenario '{scenario_name}':\n{result.stderr}")
        print('#------------------------------------------------------------------------------#')
    else:
        print(f"✅ Scenario '{scenario_name}' converted successfully.\n{result.stdout}")
        print('#------------------------------------------------------------------------------#')

def run_preprocessing_script(params, scenario_name):
    """
    Executes the preprocessing Python script specified in the YAML params file for a given scenario.

    Parameters:
        params (dict): Parameters loaded from the YAML file.
        scenario_name (str): The name of the scenario to preprocess.
    """
    # Step 1: Define paths
    script_path = os.path.join(params['Miscellaneous'], params['preprocess_data'])
    input_file = os.path.join(params['executables'], scenario_name + '_0', f"{scenario_name}_0.txt")
    output_file = os.path.join(params['executables'], scenario_name + '_0', f"{params['preprocess_data_name']}{scenario_name}_0.txt")

    # Step 2: Construct command
    command = [sys.executable, script_path, input_file, output_file]

    print(f"Running preprocessing script for scenario '{scenario_name}':")
    print(' '.join(command))

    # Step 3: Run the script
    result = subprocess.run(command, capture_output=True, text=True)

    # Step 4: Output result
    if result.returncode != 0:
        print(f"❌ Error during preprocessing of scenario '{scenario_name}':\n{result.stderr}")
        print('#------------------------------------------------------------------------------#')
    else:
        print(f"✅ Preprocessing completed for scenario '{scenario_name}':\n{result.stdout}")
        print('#------------------------------------------------------------------------------#')

def check_enviro_variables(solver_command):
    # Determine the command based on the operating system
    command = 'where' if platform.system() == 'Windows' else 'which'
    
    # Execute the appropriate command
    where_solver = subprocess.run([command, solver_command], capture_output=True, text=True)
    paths = where_solver.stdout.splitlines()
    
    if paths:  # Ensure that at least one path was found
        path_solver = paths[0]
        
        # Check if the path is already in the environment variable PATH
        if path_solver not in os.environ["PATH"]:
            # If not in PATH, add it
            os.environ["PATH"] += os.pathsep + path_solver
            print("Path added:", path_solver)
    else:
        print(f"No '{solver_command}' found on the system.")
    #

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

def main_executer(params, scenario_name):
    
    folder_scenario = os.path.join(params['executables'], scenario_name + '_0')                             
    
    # Constructing paths for the data file and the output file, adapting for file system differences
    data_file = os.path.join(folder_scenario, params['preprocess_data_name'] + scenario_name + '_0')
    output_file = os.path.join(folder_scenario, params['preprocess_data_name'] + scenario_name + '_0' + params['output_files'])
    this_case = scenario_name + '_0.txt'

    # Determining the solver based on parameters
    solver = params['solver']
    commands = []

    if solver == 'glpk':
        # Using newer GLPK options
                                   
        check_enviro_variables('glpsol')
        
        # Composing the command to solve the model with new options
        str_solve = f'glpsol -m {params["osemosys_model"]} -d {data_file}.txt --wglp {output_file}.glp --write {output_file}.sol'
        commands.append(str_solve)
        
    else:      
        # For LP models
        str_solve = f'glpsol -m {params["osemosys_model"]} -d {data_file}.txt --wlp {output_file}.lp --check'
        commands.append(str_solve)
        
        if solver == 'cbc':
            # Using CBC solver
            if os.path.exists(output_file + '.sol'):
                os.remove(output_file + '.sol')
                
            check_enviro_variables('cbc')
            
            # Composing the command for CBC solver
            str_solve = f'cbc {output_file}.lp -seconds {params["iteration_time"]} solve -solu {output_file}.sol'
            commands.append(str_solve)
            
        elif solver == 'cplex':
            # Using CPLEX solver
            if os.path.exists(output_file + '.sol'):
                os.remove(output_file + '.sol')
            
            # Number of threads cplex use
            cplex_threads = params['cplex_threads']
                                       
            check_enviro_variables('cplex')
                
            # Composing the command for CPLEX solver
            str_solve = f'cplex -c "read {output_file}.lp" "set threads {cplex_threads}" "optimize" "write {output_file}.sol"'
            commands.append(str_solve)

    for cmd in commands:
        subprocess.run(cmd, shell=True, check=True)
        
    print(f'✅ Scenario {scenario_name}_0 solve successfully.')
    print('\n#------------------------------------------------------------------------------#')

    # Paths for converting outputs
    file_path_conv_format = os.path.join(params['Miscellaneous'], params['conv_format'])
    # file_path_template = os.path.join(params['Miscellaneous'], params['templates'])
    file_path_template = os.path.join(params['A2_output_otoole'], scenario_name)
    file_path_outputs = os.path.join(folder_scenario, params['outputs'])

    # Converting outputs from .sol to csv format
    if solver == 'glpk' and params['glpk_option'] == 'new':
        str_outputs = f'otoole results {solver} csv {output_file}.sol {file_path_outputs} datafile {data_file}.txt {file_path_conv_format} --glpk_model {output_file}.glp'
        subprocess.run(str_outputs, shell=True, check=True)

    elif solver in ['cbc', 'cplex']:
                                                                                                                    
        str_outputs = f'otoole results {solver} csv {output_file}.sol {file_path_outputs} csv {file_path_template} {file_path_conv_format} 2> {output_file}.log'
        subprocess.run(str_outputs, shell=True, check=True)
    
    # Module to concatenate csvs otoole outputs
    if solver in ['glpk', 'cbc', 'cplex']:
        file_conca_csvs = get_config_main_path(os.path.abspath(''), 'config_plots')
        script_concate_csv = os.path.join(file_conca_csvs, params['concat_csvs'])
        str_otoole_concate_csv = f'python -u {script_concate_csv} {this_case} 1'  # last int is the ID tier
        subprocess.run(str_otoole_concate_csv, shell=True, check=True)
        print(f'✅ Concatenated outputs to {scenario_name}_0_Output.csv successfully.')
        print('\n#------------------------------------------------------------------------------#')

def delete_files(file, data_file, solver):
    # Delete files
    if file:
        shutil.os.remove(file)
        shutil.os.remove(data_file)
    
    # Check if the .sol file exists and is empty
    log_file = file.replace('.sol', '.log')
    if os.path.exists(log_file) and os.path.getsize(log_file) == 0:
        if os.path.exists(log_file):
            os.remove(log_file)
    
    if solver == 'glpk':
        shutil.os.remove(file.replace('sol', 'glp'))        
    else:
        shutil.os.remove(file.replace('sol', 'lp'))
    
    # Delete log files when solver is 'cplex' and del_files is True
    if solver == 'cplex':
        for filename in ['cplex.log', 'clone1.log', 'clone2.log']:
            if os.path.exists(filename):
                os.remove(filename)

def read_csv_files(input_dir):
    """Reads all CSV files in the given directory and returns a dictionary of DataFrames."""
    data_dict = {}
    for filename in os.listdir(input_dir):
        if filename.endswith(".csv"):
            file_path = os.path.join(input_dir, filename)
            df = pd.read_csv(file_path)
            key = os.path.splitext(filename)[0]
            data_dict[key] = df
    return data_dict

def generate_combined_input_file(input_folder, output_folder, scenario_name):
    """
    Reads CSVs from input_folder, filters out metadata keys, renames VALUE columns by key,
    concatenates all non-empty DataFrames, orders columns, and saves the result to a CSV file.
    """
    keys_sets_delete = ['REGION', 'YEAR', 'TECHNOLOGY', 'FUEL', 'EMISSION', 'MODE_OF_OPERATION',
                        'TIMESLICES', 'STORAGE', 'SEASON', 'DAYTYPE', 'DAILYTIMEBRACKET']

    inputs_dataframes = []
    for filename in os.listdir(input_folder):
        if not filename.endswith(".csv"):
            continue
        key = filename.replace(".csv", "")
        if key in keys_sets_delete:
            continue
        path = os.path.join(input_folder, filename)
        df = pd.read_csv(path)
        if df.empty or 'VALUE' not in df.columns:
            continue
        df = df.rename(columns={'VALUE': key})
        inputs_dataframes.append(df)

    if not inputs_dataframes:
        print("[Warning] No valid dataframes found to concatenate.")
        return None, None

    # Concatenate all non-empty dataframes
    inputs_data = pd.concat(inputs_dataframes, ignore_index=True, sort=False)

    # Reorder columns
    present_keys = [col for col in keys_sets_delete if col in inputs_data.columns]
    other_columns = sorted([col for col in inputs_data.columns if col not in present_keys])
    inputs_data = inputs_data[present_keys + other_columns]

    # Save to CSV
    os.makedirs(output_folder, exist_ok=True)
    output_path = os.path.join(output_folder, f"{scenario_name}_Input.csv")
    inputs_data.to_csv(output_path, index=False)

    print(f'✅ Concatenated inputs to {scenario_name}_Input.csv successfully.')
    print('\n#------------------------------------------------------------------------------#')

    return output_path, inputs_data.head()




########################################################################################
if __name__ == "__main__":
    # Start timer
    start1 = time.time()
    
    # Load params from YAML
    with open('MOMF_T1_AB.yaml', 'r') as f:
        params = yaml.safe_load(f)
    
    # Define source and destination base paths
    base_input_path = params['A2_output']
    template_path = os.path.join(params['Miscellaneous'], params['templates'])
    base_output_path = params['A2_output_otoole']
    
    for scenario_name in os.listdir(base_input_path):
        if scenario_name == 'Default':
            continue
        
        process_scenario_folder(
            base_input_path=base_input_path,
            template_path=template_path,
            base_output_path=base_output_path,
            scenario_name=scenario_name
        )
        
        run_otoole_conversion(
            base_output_path=base_output_path,
            scenario_name=scenario_name,
            params=params
        )
        
        run_preprocessing_script(params, scenario_name)


        input_folder = os.path.join(base_output_path, scenario_name)
        output_folder = os.path.join(params['executables'], scenario_name + '_0')
        
        # List any available files for preview (just to verify setup)
        os.makedirs(input_folder, exist_ok=True)
        os.makedirs(output_folder, exist_ok=True)
        
        # Concatenate inputs
        generate_combined_input_file(input_folder, output_folder, scenario_name + '_0')

    
        #
        main_executer(params, scenario_name)
        
        # Delete Outputs folder with otoole csvs files
        if params['del_files']:
            # Delete Outputs folder with otoole csvs files
            folder_scenario = os.path.join(params['executables'], scenario_name + '_0') 
            outputs_otoole_csvs = os.path.join(folder_scenario, params['outputs'])
            data_file = os.path.join(folder_scenario, scenario_name + '_0' + '.txt')
            sol_file = os.path.join(folder_scenario, params['preprocess_data_name'] + scenario_name + '_0' + params['output_files'] + '.sol')
            if os.path.exists(outputs_otoole_csvs):
                shutil.rmtree(outputs_otoole_csvs)
        
            # Delete glp, lp, txt and sol files
            if params['solver'] in ['glpk', 'cbc', 'cplex']:
                delete_files(sol_file, data_file, params['solver'])
            
            print(f'✅ Delete intermediate files to scenario {scenario_name}_0 successfully.')
            print('\n#------------------------------------------------------------------------------#')



    end_1 = time.time()   
    time_elapsed_1 = -start1 + end_1
    print( str( time_elapsed_1 ) + ' seconds /', str( time_elapsed_1/60 ) + ' minutes' )
    
    
    
    
    
    print('*: For all effects, we have finished the work of this script.')

            

