# -*- coding: utf-8 -*-
"""
Created on 2025

@author: ClimateLeadGroup, Andrey Salazar-Vargas
"""

import os
import pandas as pd
import yaml
import subprocess
import sys
import platform  
import shutil
import time
from datetime import date
import multiprocessing as mp
import math
from typing import List, Any
from pathlib import Path
import numpy as np

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
    scenario_exec_dir = os.path.join(HERE, params['executables'], scenario_name + '_0')
    output_file = os.path.join(scenario_exec_dir, f"{scenario_name}_0.txt")
    config_file = os.path.join(HERE, params['Miscellaneous'], params['otoole_config'])

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

    print(f"Running preprocessing script for scenario '{scenario_name}_0':")
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

def main_executer(params, scenario_name, HERE):
    
    folder_scenario = os.path.join(HERE, params['executables'], scenario_name + '_0')                             
    
    # Constructing paths for the data file and the output file, adapting for file system differences
    data_file = os.path.join(folder_scenario, params['preprocess_data_name'] + scenario_name + '_0')
    output_file = os.path.join(folder_scenario, params['preprocess_data_name'] + scenario_name + '_0' + params['output_files'])
    this_case = scenario_name + '_0.txt'

    # Determining the solver based on parameters
    solver = params['solver']
    commands = []

    if solver == 'glpk':
        if params['execute_model']:
            # Using newer GLPK options
                                       
            check_enviro_variables('glpsol')
            
            # Composing the command to solve the model with new options
            str_solve = f'glpsol -m {params["osemosys_model"]} -d {data_file}.txt --wglp {output_file}.glp --write {output_file}.sol'
            commands.append(str_solve)
        
    else:
        if params['create_matrix']:
            # For LP models
            str_solve = f'glpsol -m {params["osemosys_model"]} -d {data_file}.txt --wlp {output_file}.lp --check'
            commands.append(str_solve)
        
        if solver == 'cbc':
            # Using CBC solver
            if params['execute_model']:
                if os.path.exists(output_file + '.sol'):
                    os.remove(output_file + '.sol')
                
                check_enviro_variables('cbc')
                
                # Composing the command for CBC solver
                str_solve = f'cbc {output_file}.lp -seconds {params["iteration_time"]} solve -solu {output_file}.sol'
                commands.append(str_solve)
            
        elif solver == 'cplex':
            # Using CPLEX solver
            if params['execute_model']:
                if os.path.exists(output_file + '.sol'):
                    os.remove(output_file + '.sol')
            
                # Number of threads cplex use
                cplex_threads = params['cplex_threads']
                                           
                check_enviro_variables('cplex')
                    
                # Composing the command for CPLEX solver
                str_solve = f'cplex -c "read {output_file}.lp" "set threads {cplex_threads}" "optimize" "write {output_file}.sol"'
                commands.append(str_solve)
    if params['execute_model'] or params['create_matrix']:
        for cmd in commands:
            subprocess.run(cmd, shell=True, check=True)
        
    print(f'✅ Scenario {scenario_name}_0 solve successfully.')
    print('\n#------------------------------------------------------------------------------#')

    # Paths for converting outputs
    file_path_conv_format = os.path.join(HERE, params['Miscellaneous'], params['conv_format'])
    # file_path_template = os.path.join(params['Miscellaneous'], params['templates'])
    file_path_template = os.path.join(HERE, params['A2_output_otoole'], scenario_name)
    file_path_outputs = os.path.join(folder_scenario, params['outputs'])

    # Converting outputs from .sol to csv format
    if solver == 'glpk' and params['glpk_option'] == 'new':
        str_outputs = f'otoole results {solver} csv {output_file}.sol {file_path_outputs} datafile {data_file}.txt {file_path_conv_format} --glpk_model {output_file}.glp'
        if params['execute_model']:
            subprocess.run(str_outputs, shell=True, check=True)

    elif solver in ['cbc', 'cplex']:
                                                                                                                    
        str_outputs = f'otoole results {solver} csv {output_file}.sol {file_path_outputs} csv {file_path_template} {file_path_conv_format} 2> {output_file}.log'
        if params['execute_model']:
            subprocess.run(str_outputs, shell=True, check=True)
    
    # Module to concatenate csvs otoole outputs
    if solver in ['glpk', 'cbc', 'cplex']:
        file_conca_csvs = get_config_main_path(os.path.abspath(''), params['concatenate_folder'])
        script_concate_csv = os.path.join(file_conca_csvs, params['concat_csvs'])
        str_otoole_concate_csv = f'python -u {script_concate_csv} {file_path_outputs} {output_file}'  # last int is the ID tier
        if params['concat_otoole_csv']:
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
                        'TIMESLICE', 'STORAGE', 'SEASON', 'DAYTYPE', 'DAILYTIMEBRACKET']

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



def concatenate_all_scenarios(HERE, params):
    """
    Iterates over all scenario folders in `base_input_path` (excluding 'Default'),
    reads *_Input.csv and *_Output.csv files, adds scenario metadata columns, concatenates
    them into single CSV files for inputs, outputs y combined, y devuelve sus rutas.

    Args:
        params (dict):
          - executables (str): Path to the base directory containing scenario folders.
          - prefix_final_files (str): Carpeta/ruta donde guardar los resultados.
          - inputs_file (str): Nombre base para el CSV de inputs.
          - outputs_file (str): Nombre base para el CSV de outputs.
          - combined_file (str, opcional): Nombre base para el CSV combinado inputs+outputs.
    Returns:
        tuple: (input_csv_path, output_csv_path, combined_csv_path)
    """
    # Columnas de metadatos que movemos al frente
    keys_sets_delete = [
        'REGION','YEAR','TECHNOLOGY','FUEL','EMISSION','MODE_OF_OPERATION',
        'TIMESLICE','STORAGE','SEASON','DAYTYPE','DAILYTIMEBRACKET'
    ]

    combined_inputs = []
    combined_outputs = []
    combined_inputs_outputs = []
    base_input_path = params['executables']

    for scenario_future_name in os.listdir(base_input_path):
        if scenario_future_name.lower() in ['default', '__pycache__', 'local_dataset_creator_0.py']:
            continue

        scenario_path = os.path.join(HERE, base_input_path, scenario_future_name)
        parts = scenario_future_name.rsplit("_", 1)
        scenario = parts[0]
        future = parts[1]

        input_file = os.path.join(scenario_path, f"{scenario_future_name}_Input.csv")
        output_file = os.path.join(scenario_path, f"Pre_processed_{scenario_future_name}_Output.csv")

        if os.path.exists(input_file):
            df_in = pd.read_csv(input_file, low_memory=False)
            df_in.insert(0, "Future", future)
            df_in.insert(1, "Scenario", scenario)
            combined_inputs.append(df_in)
            combined_inputs_outputs.append(df_in)

        if os.path.exists(output_file):
            df_out = pd.read_csv(output_file, low_memory=False)
            df_out.insert(0, "Future", future)
            df_out.insert(1, "Scenario", scenario)
            combined_outputs.append(df_out)
            combined_inputs_outputs.append(df_out)

    # Concatenate inputs y outputs por separado
    df_inputs_all = pd.concat(combined_inputs, ignore_index=True) if combined_inputs else pd.DataFrame()
    df_outputs_all = pd.concat(combined_outputs, ignore_index=True) if combined_outputs else pd.DataFrame()
    # df_inputs_outputs_all = pd.concat(combined_inputs_outputs, ignore_index=True) if combined_inputs_outputs else pd.DataFrame()
    # df_list = []
    # df_list.append(combined_inputs)
    # df_list.append(combined_outputs)
    df_inputs_outputs_all = pd.concat([df_inputs_all,df_outputs_all], ignore_index=True, sort=False)
    

    # Función para reordenar columnas: metadata first, then alfabetico
    def reorder_columns(df):
        front = ['Future','Scenario'] + [c for c in keys_sets_delete if c in df.columns]
        rest = sorted([c for c in df.columns if c not in front])
        return df[front + rest]

    today = date.today().isoformat()  # 'YYYY-MM-DD'

    # 1) Guardar inputs
    if not df_inputs_all.empty:
        df_inputs_all = reorder_columns(df_inputs_all)
        path_in = os.path.join(HERE,params['prefix_final_files'] + params['inputs_file'])
        df_inputs_all.to_csv(path_in, index=False)
        dated = path_in.replace('.csv', f'_{today}.csv')
        df_inputs_all.to_csv(dated, index=False)
    else:
        path_in = None

    # 2) Guardar outputs
    if not df_outputs_all.empty:
        df_outputs_all = reorder_columns(df_outputs_all)
        path_out = os.path.join(HERE,params['prefix_final_files'] + params['outputs_file'])
        df_outputs_all.to_csv(path_out, index=False)
        dated = path_out.replace('.csv', f'_{today}.csv')
        df_outputs_all.to_csv(dated, index=False)
    else:
        path_out = None

    # 3) Nuevamente, combinar ambos DataFrames en uno solo y guardarlo
    combined_name = params.get('combined_file', 'Combined_Inputs_Outputs.csv')
    if not df_inputs_outputs_all.empty and not df_outputs_all.empty:
        # df_combined = pd.concat([df_inputs_all, df_outputs_all],
        #                         ignore_index=True, sort=False)
        df_combined = reorder_columns(df_inputs_outputs_all)
        
        
        #########################################################################################
        df = df_combined.copy()  # para evitar vistas
        df['AccumulatedTotalAnnualMinCapacityInvestment'] = df['TotalAnnualMinCapacityInvestment']
        
        # 2) Determina el rango de años dinámicamente
        years = sorted(df['YEAR'].dropna().unique())
        period_start = years[0]   # p.ej. 2021
        period_end   = years[-1]  # p.ej. 2050
        
        # 3) Inicializa el acumulador
        acc = 0
        
        # 4) Recorre fila a fila SIN agrupar ni filtrar
        for idx in df.index:
            year = df.at[idx, 'YEAR']
            val  = df.at[idx, 'AccumulatedTotalAnnualMinCapacityInvestment']
            
            # Cuando llegue al año inicial, reinicias el acumulador
            if year == period_start:
                acc = val
                df.at[idx, 'AccumulatedTotalAnnualMinCapacityInvestment'] = val
            else:
                # En años posteriores, sumas el valor actual al acumulado
                acc = acc + val
                df.at[idx, 'AccumulatedTotalAnnualMinCapacityInvestment'] = acc
        df_combined = df


        #########################################################################################
        
        
        path_comb = os.path.join(HERE,params['prefix_final_files'] + combined_name)
        df_combined.to_csv(path_comb, index=False)
        dated = path_comb.replace('.csv', f'_{today}.csv')
        df_combined.to_csv(dated, index=False)
    else:
        path_comb = None

    return path_in, path_out, path_comb






def chunk_scenarios(
    scenarios: List[Any],
    max_x_per_iter: int,
) -> List[List[Any]]:
    """
    Split the input list ``scenarios`` into chunks of size ``max_x_per_iter``.

    Parameters
    ----------
    scenarios : List[Any]
        The list that holds all scenario values.
    max_x_per_iter : int
        Maximum number of elements allowed in each chunk.

    Returns
    -------
    List[List[Any]]
        A list where each element is a sub-list of ``scenarios`` with length
        up to ``max_x_per_iter``.
    """
    if max_x_per_iter <= 0:
        raise ValueError("max_x_per_iter must be a positive integer")

    # Build the chunks using slicing in a comprehension
    scenarios_list_max_per_iter: List[List[Any]] = [
        scenarios[i : i + max_x_per_iter]  # noqa: E203 (spacing around :)
        for i in range(0, len(scenarios), max_x_per_iter)
    ]
    return scenarios_list_max_per_iter

########################################################################################
if __name__ == "__main__":
    # Start timer
    start1 = time.time()
    
    # Carpeta donde vive este script: .../relac_tx/t1_confection
    global HERE
    HERE = Path(__file__).resolve().parent
    
    
    # (Opcional) Cambiar CWD a la carpeta del script
    if Path.cwd() != HERE:
        os.chdir(HERE)
        print(f"[INFO] Working dir -> {HERE}")
        
    # Load params from YAML
    with open('MOMF_T1_AB.yaml', 'r') as f:
        params = yaml.safe_load(f)
        
    # Load params from YAML
    with open('MOMF_T1_A.yaml', 'r') as f:
        params_A2 = yaml.safe_load(f)
    
    # Define source and destination base paths
    base_input_path = os.path.join(HERE, params['A2_output'])
    template_path = os.path.join(HERE, params['Miscellaneous'], params['templates'])
    base_output_path = os.path.join(HERE, params['A2_output_otoole'])
    
    scenarios=os.listdir(base_input_path)
    try:
        scenarios.remove('Default')
    except ValueError:
        pass
    
    if params['only_main_scenario']:
        scenarios = []
        scenarios.append(params_A2['xtra_scen']['Main_Scenario'])
    
    ###############################################################################################
    # Write txt model
    for scenario_name in scenarios:
         
        if params['write_txt_model']:
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


        input_folder = os.path.join(HERE, base_output_path, scenario_name)
        output_folder = os.path.join(HERE, params['executables'], scenario_name + '_0')
        
        # List any available files for preview (just to verify setup)
        os.makedirs(input_folder, exist_ok=True)
        os.makedirs(output_folder, exist_ok=True)
        
        # Concatenate inputs
        generate_combined_input_file(input_folder, output_folder, scenario_name + '_0')

        #
    ###############################################################################################
        
        
        
    ###############################################################################################
    # Execute txt model
    if params['execute_model'] or params['create_matrix']:
        if params['parallel']:
            print('Entered Parallelization of model execution')
            max_x_per_iter = params['max_x_per_iter'] # FLAG: This is an input
            scenarios_list_max_per_iter = chunk_scenarios(scenarios, max_x_per_iter)
            #
            for scens_list in scenarios_list_max_per_iter:
                processes = []
                for scenario_name in scens_list:
                    p = mp.Process(target=main_executer, args=(params, scenario_name, HERE) )
                    processes.append(p)
                    p.start()
                #
                for process in processes:
                    process.join()
            
        # This is for the linear version
        else:
            print('Started Linear Runs')
            for scenario_num in scenarios:
                main_executer(params, scenario_num, HERE)
    
    ###############################################################################################
    # Delete files
    for scenario_name in scenarios:        
        # Delete Outputs folder with otoole csvs files
        if params['del_files']:
            # Delete Outputs folder with otoole csvs files
            folder_scenario = os.path.join(HERE, params['executables'], scenario_name + '_0') 
            outputs_otoole_csvs = os.path.join(HERE, folder_scenario, params['outputs'])
            data_file = os.path.join(HERE, folder_scenario, scenario_name + '_0' + '.txt')
            sol_file = os.path.join(HERE, folder_scenario, params['preprocess_data_name'] + scenario_name + '_0' + params['output_files'] + '.sol')
            if os.path.exists(outputs_otoole_csvs):
                shutil.rmtree(outputs_otoole_csvs)
        
            # Delete glp, lp, txt and sol files
            if params['solver'] in ['glpk', 'cbc', 'cplex']:
                delete_files(sol_file, data_file, params['solver'])
            
            print(f'✅ Delete intermediate files to scenario {scenario_name}_0 successfully.')
            print('\n#------------------------------------------------------------------------------#')

    ###############################################################################################

    end_1 = time.time()   
    time_elapsed_1 = -start1 + end_1
    print( str( time_elapsed_1 ) + ' seconds /', str( time_elapsed_1/60 ) + ' minutes' )

    start2 = time.time()
    
    ###############################################################################################
    # Concatenate inputs and outputs
    if params['concat_scenarios_csv']:
        input_output_path, output_output_path, combined_output_path = concatenate_all_scenarios(HERE,params)
        print(f'✅ Concatenate inputs and outputs for all scenarios successfully.')
        print(f'The name files are: ({input_output_path}), ({output_output_path}) and ({combined_output_path})')
    ###############################################################################################
    
    
    
    
    # # 1. Carga los dataframes desde los CSV
    # df_inputs_all = pd.read_csv('REALC_TX_Inputs.csv', low_memory=False)
    # df_outputs_all = pd.read_csv('REALC_TX_Outputs.csv', low_memory=False)
    
    # # 2. Concaténalos verticalmente (uno debajo del otro)
    # df_combined = pd.concat([df_inputs_all, df_outputs_all], ignore_index=True, sort=False)
    
    # # 3. (Opcional) Reordena columnas si lo deseas,
    # #    por ejemplo, poniendo 'Scenario' y 'Future' al frente
    # cols_front = ['Scenario', 'Future']
    # other_cols = [c for c in df_combined.columns if c not in cols_front]
    # df_combined = df_combined[cols_front + other_cols]
    
    # # 4. Guarda el CSV combinado
    # today = date.today().isoformat()  # e.g. '2025-07-14'
    # combined_filename = f"{params['prefix_final_files']}Combined_Inputs_Outputs_{today}.csv"
    # df_combined.to_csv(combined_filename, index=False)
    
    # print(f"Archivo combinado guardado en: {combined_filename}")
    ###############################################################################################
    
    
    end_2 = time.time()   
    time_elapsed_2 = -start2 + end_2
    print( str( time_elapsed_2 ) + ' seconds /', str( time_elapsed_2/60 ) + ' minutes' )
    print('\n#------------------------------------------------------------------------------#')
    
    time_elapsed_3 = -start1 + end_2
    print( str( time_elapsed_3 ) + ' seconds /', str( time_elapsed_3/60 ) + ' minutes' )
    print('*: For all effects, we have finished the work of this script.')

            

