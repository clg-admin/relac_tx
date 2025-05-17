# -*- coding: utf-8 -*-
"""
Created on Mon May 13 09:57:06 2024

@author: andreysava19
"""

# Code to verify missing technologies in various Excel files across different scenarios

import pandas as pd
import pickle
import os
import re
import yaml
import sys


# Note: dor this script children and daughters are the same


###############################################  Functions  ################################################################
# Function to load data and check for missing technologies
def find_missing_technologies(scenario, file_name):
    file_path = os.path.join(base_path, scenario, file_name)
    data = pd.read_csv(file_path)
    existing_technologies = set(data['TECHNOLOGY'].unique())
    missing_technologies = all_technologies - existing_technologies
    return missing_technologies

def check_tech_presence(missing_info, fleet_groups):
    for scenario, files in missing_info.items():
        for file_name, missing_techs in files.items():
            for tech_mother, tech_children in fleet_groups.items():
                if tech_mother not in missing_techs:  # Check if the mother technology is present
                    missing_children = [child for child in tech_children if child in missing_techs]
                    if missing_children:  # If there are missing child technologies, print information
                        print(f"Scenario: {scenario} \nParameter: {file_name}")
                        print(f"Mother technology present: {tech_mother}")
                        print(f"Missing child technologies: {missing_children}")
                        print("-------------------------------------------------")
                        
def check_mother_absence(missing_info, fleet_groups):
    for scenario, files in missing_info.items():
        for file_name, missing_techs in files.items():
            for tech_mother, tech_children in fleet_groups.items():
                present_children = [child for child in tech_children if child not in missing_techs]
                if present_children and tech_mother in missing_techs:  # If the mother is absent but children are not
                    print(f"Scenario: {scenario} \nParameter: {file_name}")
                    print(f"Child technologies present: {present_children}")
                    print(f"Mother technology absent: {tech_mother}")
                    print("-------------------------------------------------")


def read_parameters(file_path, parameter_name):
    data = {}  # This dictionary will store the extracted data
    years = []  # This list will store the years

    with open(file_path, 'r') as file:
        lines = file.readlines()

    # Search for the parameter definition
    start_index = -1
    for i, line in enumerate(lines):
        if f"param {parameter_name} default" in line:
            # Find the line containing the years, assumed to be two lines after the definition
            years_line = lines[i + 2]
            # Clean the years line and convert to a list of integers
            years = [int(year) for year in re.findall(r'\d{4}', years_line)]
            start_index = i + 3
            break
    
    if start_index == -1:
        print(f"The restriction {parameter_name} was not found.")
        return pd.DataFrame()

    # Search and capture values for technologies until a ';' is found
    for line in lines[start_index:]:
        if line.strip() == ';':
            break
        elements = line.split()
        technology = elements[0]
        values = list(map(float, elements[1:]))  # Convert the values to float
        data[technology] = values

    # Convert the dictionary to a pandas DataFrame for easier manipulation
    return pd.DataFrame(data, index=years).T  # Transpose so that technologies are rows and years are columns

def read_parameters_variant(file_path, parameter_name):
    data = {}  # Dictionary to store extracted data
    years = []  # List to store the years

    with open(file_path, 'r') as file:
        lines = file.readlines()

    # Search for the parameter definition
    start_index = -1
    for i, line in enumerate(lines):
        if f"param {parameter_name} default" in line:
            start_index = i + 1
            break
    
    if start_index == -1:
        print(f"The parameter {parameter_name} definition was not found.")
        return pd.DataFrame()

    # Read blocks of data
    i = start_index
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith('['):
            # Extract technologies and other identifiers
            identifiers = line.strip('[]').split(',')
            tech_identifier = f"{identifiers[1]}" # / {identifiers[-3]}"
            
            # Extract years from the next line
            years_line = lines[i + 1].strip().split()
            # Remove ':=', if present
            if ':=' in years_line:
                years_line.remove(':=')
            years = [int(year) for year in years_line if year.isdigit()]
            
            # Extract values from the line after the years
            values_line = lines[i + 2].strip().split()
            # Stop processing if ';' is found
            # if ';' in values_line:
            #     values_line.remove(';')
            
            # Ignore the first value (always 1) and convert the rest to float
            values = [float(value) for value in values_line[1:] if re.match(r'^-?\d+(?:\.\d+)?$', value)]
            
            # Ensure the lengths match before storing
            if len(values) == len(years):
                data[tech_identifier] = values
            else:
                print(f"Mismatch in lengths for {tech_identifier}: {len(values)} values, {len(years)} years")
            
        # Stop processing if ';' is found
        if ';' in line:
            break

            # Move to the next block
            i += 3
        else:
            i += 1

    # Convert the dictionary to a pandas DataFrame
    return pd.DataFrame(data, index=years).T  # Transpose so that technologies are rows and years are columns

def compare_mother_daughters(parameter, results, fleet_groups, output_filename, i, num):
    comparison_results = {}  # This dictionary will store the comparison results
    count = 0

    if i == 0 :
        option = 'w'
    else:
        option = 'a'
    # Open the file to write the results
    with open(output_filename, option) as file:
        if option == 'w':
            file.write(f"######################################################## TEST {num} ########################################################\n")
            file.write("Problem if sum of daughters technologies are greater than mother technologies for the same parameter,\nonly for transport sector.\n")
        # Iterate over each mother technology in the group dictionary
        for tech_mother, children in fleet_groups.items():
            if tech_mother in results.index:
                comparison_results[tech_mother] = {}
                children_sum = pd.Series([0.0] * len(results.columns), index=results.columns)
                children_without_values = []

                for child in children:
                    if child in results.index:
                        # Check if the child technology has all values as zero
                        if results.loc[child].sum() == 0:
                            children_without_values.append(child)
                        children_sum += results.loc[child]
                    else:
                        # If the child technology is not in the index, it is considered to have no values
                        children_without_values.append(child)

                mother_values = results.loc[tech_mother]
                for year in results.columns:
                    if children_sum[year] > mother_values[year]:
                        comparison = "greater"
                    elif children_sum[year] < mother_values[year]:
                        comparison = "less"
                    else:
                        comparison = "equal"
                    
                    if comparison == "greater":
                        comparison_results[tech_mother][year] = comparison + '      mother value: ' + str(mother_values[year]) + ',   children value: ' + str(children_sum[year])

                # Only write results if there are any
                if comparison_results[tech_mother]:
                    if count == 0:
                        file.write(f"\n\n\nResults for {parameter}:\n")
                    # Write the results for the mother technology to the file
                    file.write(f"   Results for {tech_mother}:\n")
                    if children_without_values:
                        file.write(f"     Children without values: {children_without_values}\n") 
                    for year, result in comparison_results[tech_mother].items():
                        file.write(f"     Check   {year}: {result}\n")
                    file.write("-------------------------------------------------\n")
        file.write("-----------------------------------------------------------------------------------------------\n")
        count += 1

    return comparison_results


def check_decreasing_values(data_frame, parameter_name, output_filename, num):
    """ Check if the values for each technology are strictly decreasing over the years. """
    # Dictionary to store technologies with non-decreasing issues
    non_decreasing_issues = {}

    # Iterate through each technology
    for tech, values in data_frame.iterrows():
        # Use pd.to_numeric to convert values to float and handle non-numeric data gracefully
        numeric_values = pd.to_numeric(values, errors='coerce')
        # Drop NaN values which might occur due to conversion or invalid data
        numeric_values = numeric_values.dropna()
        # Check if the series is strictly decreasing
        if not all(numeric_values.iloc[i] > numeric_values.iloc[i + 1] for i in range(len(numeric_values) - 1)):
            non_decreasing_issues[tech] = numeric_values.tolist()

    # Open output file for appending results
    with open(output_filename, 'a') as file:
        file.write(f"\n\n\n\n\n\n######################################################## TEST {num} ########################################################\n")
        file.write(f"Problem if technologies have strictly decreasing values for {parameter_name}, by available technologies.\n")
        # Print warning messages if any non-decreasing issues are found
        if non_decreasing_issues:
            file.write("Warning: The following technologies do not have strictly decreasing values:\n")
            for tech, vals in non_decreasing_issues.items():
                file.write(f"Check   {tech}: {vals}\n")
        else:
            file.write("All values are strictly decreasing for all technologies.\n")

def check_any_decreasing_values(data_frame, parameter_name, output_filename, num):
    """ Check if there are any decreasing values from one year to the next across each technology. """
    # Dictionary to store technologies with any decreasing values found
    decreases_found = {}

    # Iterate through each technology
    for tech, values in data_frame.iterrows():
        # Convert values to numeric, handling non-numeric data gracefully
        numeric_values = pd.to_numeric(values, errors='coerce')
        # Drop NaN values which might occur due to conversion or invalid data
        numeric_values = numeric_values.dropna()
        # Initialize list to store decreases
        decreases = []
        # Iterate through each pair of consecutive years
        years = numeric_values.index
        for i in range(len(years) - 1):
            if numeric_values[years[i]] > numeric_values[years[i+1]]:
                decreases.append((years[i], numeric_values[years[i]], years[i+1], numeric_values[years[i+1]]))

        if decreases:
            decreases_found[tech] = decreases

    # Open output file for appending results
    with open(output_filename, 'a') as file:
        file.write(f"\n\n\n\n\n\n######################################################## TEST {num} ########################################################\n")
        file.write(f"Problem with decreasing values of {parameter_name} by available technologies, by available technologies.\n")
        # Print results if any decreases are found
        if decreases_found:
            for tech, dec in decreases_found.items():
                file.write(f"{tech} has decreases at the following points:\n")
                for decrease in dec:
                    file.write(f"  Check   from Year {decrease[0]}: {decrease[1]} to Year {decrease[2]}: {decrease[3]}\n")
                file.write("\n-------------------------------------------------\n")
        else:
            file.write("No decreases found across any technology.\n")
            file.write("\n-------------------------------------------------\n")

def compare_techs(upper_techs, lower_techs, output_filename, param_names, num):
    # Intersection of indices to ensure operations only on common rows
    common_indices = upper_techs.index.intersection(lower_techs.index)

    # Filter DataFrames to only include common rows
    filtered_upper = upper_techs.loc[common_indices]
    filtered_lower = lower_techs.loc[common_indices]

    # Perform the subtraction
    diff_df = filtered_upper - filtered_lower
    
    # Open output file for appending results
    with open(output_filename, 'a') as file:
        file.write(f"\n\n\n\n\n\n\n######################################################## TEST {num} ########################################################\n")
        file.write(f"\nProblem if values of {param_names[0]} are greater than {param_names[1]},\nby available technologies.\n")
        # Iterate over each row to print the results
        for index, row in diff_df.iterrows():
            negative_diff_exists = False  # Flag to check for negative differences
            result_lines = []  # Store lines temporarily
            for year, value in row.items():
                if value < 0:  # Only print if the difference is negative
                    result_lines.append(f"\n  Check   negative difference in the year {year}: {value}")
                    negative_diff_exists = True
            if negative_diff_exists:
                file.write(f"\nResults for technology: {index}")  # Initial index print
                for line in result_lines:
                    file.write(line)
                file.write("\n-------------------------------------------------\n")

def check_demand_vs_capacity(df_mode_broad, parameter, spec_an_dem_techs, oar_techs, output_filename, param_names, num):
    """
    Function to check if param_names[0] is less than param_names[1] x param_names[2]
    for each grandmother technology.
    """
    # Open output file for appending results
    with open(output_filename, 'a') as file:
        file.write(f"\n\n\n\n\n\n\n######################################################## TEST {num} ########################################################\n")
        file.write(f"\nProblem if values of {param_names[0]}(fuels/grandma) are greater than\n{param_names[1]}(mothers) x {param_names[2]}(mothers), only for transport sector.\n")
        
        # Calculate the equation for each grandmother technology
        for grandmother in df_mode_broad.columns[1:]:
            mother_techs = df_mode_broad[df_mode_broad[grandmother] == 'x']['Techs/Demand']
        
            for year in parameter.columns:
                sum_products = 0
                for mother in mother_techs:
                    if mother in parameter.index and mother in oar_techs.index:
                        max_cap = parameter.loc[mother, year]
                        output_ratio = oar_techs.loc[mother, year]
                        sum_products += max_cap * output_ratio
        
                specified_demand = spec_an_dem_techs.loc[grandmother, year]
                # Check if the sum of products is greater than or equal to the specified demand
                if sum_products <= specified_demand:
                    file.write(f" Check   for {grandmother} in year {year}, the condition is NOT satisfied.\n")

def check_demand_vs_capacity_variant(parameter, parameter_2, parameter_3, output_filename, param_names, num, parameter_4=None):
    """
    Function to check if param_names[0] is less than param_names[1] x param_names[2]
    for each common technology present in the indices of parameter, parameter_2, and parameter_3.
    """
    # Find common indices (technologies)
    common_techs = parameter.index.intersection(parameter_2.index).intersection(parameter_3.index)

    # Open output file for appending results
    with open(output_filename, 'a') as file:
        file.write(f"\n\n\n\n\n\n\n######################################################## TEST {num} ########################################################\n")
        if parameter_4 is None:
            file.write(f"\nProblem if values of {param_names[2]} are less than\n{param_names[0]} x {param_names[1]} x 31.356, by available technologies.\n")
        else:
            file.write(f"\nProblem if values of {param_names[2]} are less than\n{param_names[0]} x {param_names[1]} x {param_names[3]} x 31.356, by available technologies.\n")
        # Calculate the equation for each common technology
        
        for tech in common_techs:
            for year in parameter.columns:
                param_1 = parameter.loc[tech, year]
                param_2 = parameter_2.loc[tech, year]
                param_3 = parameter_3.loc[tech, year]
                if parameter_4 is None:
                    product = param_1 * param_2 * 31.356
                else:
                    param_4 = parameter_4.loc[tech, year]
                    product = param_1 * param_2 * param_4 * 31.356
                
                # Check if the product is greater than or equal to the specified demand
                if not product >= param_3:
                    file.write(f"Check   for {tech} in year {year}, the condition is NOT satisfied.\n")

def check_maxcapacity_vs_lowerlimit(param_1, param_2, output_filename, params_names, num):
    common_techs = param_1.index.intersection(param_2.index)
    common_techs_tr = [tech for tech in common_techs if 'PP' not in tech[:2]]
    years = param_1.columns.tolist()
    # Open output file for appending results
    with open(output_filename, 'a') as file:
        file.write(f"\n\n\n\n\n\n\n######################################################## TEST {num} ########################################################\n")
        file.write(f'\nProblem if {params_names[0]} is less than {params_names[1]}, only for transport sector.\n')
        for tech in common_techs_tr:
            # for year in parameter.columns:
            param_maxcap = param_1.loc[tech].tolist()
            param_lowerlimit = param_2.loc[tech].tolist()
           
            diff_list = [a - b for a, b in zip(param_maxcap, param_lowerlimit)]
            # Check if there are any negative values in the list
            has_negative = any(diff < 0 for diff in diff_list)
            
            
            if has_negative:
                file.write(f'\n   Results for {tech}:\n')
                
                for dif in range(len(diff_list)):
                    if diff_list[dif]<0:
                        file.write('Check   ' + str(years[dif]) + ': greater   MaxCapacity value: ' + str(param_maxcap[dif]) + ',   LowerLimit value: ' + str(param_lowerlimit[dif]) + '\n')
                file.write('-------------------------------------------------\n')

def compare_mother_daughters_variant(param_1, param_2, fleet_groups, output_filename, params_names, num):
    comparison_results = {}  # This dictionary will store the comparison results

    # Open the file to write the results
    with open(output_filename, 'a') as file:
        file.write(f"\n\n\n\n\n\n\n######################################################## TEST {num} ########################################################\n")
        file.write(f"\nProblem if sum of {params_names[0]}(daughters) are less than {params_names[1]}(mother),\nonly for transport sector.\n")
        # Iterate over each mother technology in the group dictionary
        for tech_mother, children in fleet_groups.items():
            if tech_mother in param_1.index:
                comparison_results[tech_mother] = {}
                children_sum = pd.Series([0.0] * len(param_1.columns), index=param_1.columns)

                for child in children:
                    if child in param_2.index:
                        children_sum += param_2.loc[child]

                mother_values = param_1.loc[tech_mother]
                for year in param_1.columns:
                    if children_sum[year] > mother_values[year]:
                        comparison = "greater"
                    elif children_sum[year] < mother_values[year]:
                        comparison = "less"
                        comparison_2 = - mother_values[year] + children_sum[year]
                    else:
                        comparison = "equal"
                    
                    if comparison == "less":
                        comparison_results[tech_mother][year] = str(comparison_2) + f'      {params_names[1]} value: ' + str(mother_values[year]) + f',   Sum {params_names[0]} value: ' + str(children_sum[year])

                # Only write results if there are any
                if comparison_results[tech_mother]:
                    # Write the results for the mother technology to the file
                    file.write(f"   Results for {tech_mother}:\n")            
                    for year, result in comparison_results[tech_mother].items():
                        file.write(f"     Check   {year}: {result}\n")
                    file.write("-------------------------------------------------\n")

def compare_mother_daughters_variant_2(param_1, param_2, fleet_groups, output_filename, params_names, num):
    comparison_results = {}  # This dictionary will store the comparison results

    # Open the file to write the results
    with open(output_filename, 'a') as file:
        file.write(f"\n\n\n\n\n\n\n######################################################## TEST {num} ########################################################\n")
        file.write(f"\nProblem if sum of {params_names[0]}(daughters) are greater than {params_names[1]}(mother),\nonly for afolu sector.\n")
        # Iterate over each mother technology in the group dictionary
        for tech_mother, children in fleet_groups.items():
            if tech_mother in param_1.index:
                comparison_results[tech_mother] = {}
                children_sum = pd.Series([0.0] * len(param_1.columns), index=param_1.columns)

                for child in children:
                    if child in param_2.index:
                        children_sum += param_2.loc[child]

                mother_values = param_1.loc[tech_mother]
                for year in param_1.columns:
                    if children_sum[year] > mother_values[year]:
                        comparison = "greater"
                        comparison_2 = - mother_values[year] + children_sum[year]
                    elif children_sum[year] < mother_values[year]:
                        comparison = "less"
                        comparison_2 = - mother_values[year] + children_sum[year]
                    else:
                        comparison = "equal"
                    
                    if comparison == "greater":
                        comparison_results[tech_mother][year] = str(comparison_2) + f'      {params_names[1]} value: ' + str(mother_values[year]) + f',   Sum {params_names[0]} value: ' + str(children_sum[year])

                # Only write results if there are any
                if comparison_results[tech_mother]:
                    # Write the results for the mother technology to the file
                    file.write(f"   Results for {tech_mother}:\n")            
                    for year, result in comparison_results[tech_mother].items():
                        file.write(f"     Check   {year}: {result}\n")
                    file.write("-------------------------------------------------\n")

# Function to search the path of yaml file
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
    appended_path = os.path.join(base_path, base_folder) + os.sep
    
    return appended_path

# Function to load yaml file and assing year of discount rate to some values
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

############################################################################################################################

if __name__ == '__main__':
    
    # Take the case from the model (B1 or experiment_manager)
    main_path = sys.argv
    initial_path = main_path[1]
    tier = main_path[2]
    
    #------------------------------------------------------------------------------------------------------------------------------#
    # Configuration parameters
    # Select tier
    # tier = 1 # 1 for B1 and 3 for RDM
    base_case_RDM = False # Only for RDM, if you use specific case or tier = 1  put in False
    
    # For specific case
    specific_case = False # if you want other file outside of the Executables or futures folders, if you use tier = 1 put in False
    specific_case_path_txt = 'LTS_5.txt' # Put the path of the input file there
    specific_case_pickle_path = 'A-O_Fleet_Groups.pickle' # Put the path of this pickle file
    specific_case_file_path_modes_transp = 'A-I_Classifier_Modes_Transport.xlsx' # Put the path of this excel file
    
    
    # Change for each model to use, this list has every transport technology
    transport_techs = [
    "TRAUTDSL", "TRAUTELE", "TRAUTGSL", "TRAUTHG", "TRAUTLPG", 
    "TRBPUDSL", "TRBPUELE", "TRBPUGSL", "TRBPULPG", "TRBTURDSL", 
    "TRBTURELE", "TRBTURGSL", "TRBTURLPG", "TRMBSDSL", "TRMBSELE", 
    "TRMBSGSL", "TRMBSLPG", "TRMOTELE", "TRMOTGSL", "TRSUVDSL", 
    "TRSUVELE", "TRSUVGSL", "TRSUVLPG", "TRSUVNGS", "TRTAXELE", 
    "TRTAXGSL", "TRTAXLPG", "TRYLFDSL", "TRYLFELE", "TRYLFGSL", 
    "TRYLFLPG", "TRYTKDSL", "TRYTKELE", "TRYTKGSL", "TRYTKHYD", 
    "TRYTKLPG", "Techs_Auto", "Techs_Buses_Micro", "Techs_Buses_Pub", 
    "Techs_Buses_Tur", "Techs_He_Freight", "Techs_Li_Freight", 
    "Techs_Motos", "Techs_SUV", "Techs_Taxi", "Techs_Telef", 
    "Techs_Trains", "Techs_Trains_Freight"
    ]
    #------------------------------------------------------------------------------------------------------------------------------#
    
    # Read yaml file
    file_config_address = get_config_main_path(os.path.abspath(''), 'config_main_files')
    params = load_and_process_yaml(os.path.join(file_config_address, 'MOMF_B1_exp_manager.yaml'))
  
    if specific_case:
        pickle_path = specific_case_pickle_path
    elif tier == str(1):
        path_fleet_group = os.path.join('A1_outputs', 'A-O_Fleet_Groups.pickle')
        pickle_path = os.path.join(initial_path, path_fleet_group)
    else:
        path_fleet_group = os.path.join('0_From_Confection', 'A-O_Fleet_Groups.pickle')
        pickle_path = os.path.join(initial_path, path_fleet_group)
        
    # Load mother and child technologies from the .pickle file
    with open(pickle_path, 'rb') as file:
        fleet_groups = pickle.load(file)
    
    # List of all technology names (mother and daughters)
    all_technologies = set(fleet_groups.keys())  # Include mother technologies
    for techs in fleet_groups.values():
        all_technologies.update(techs)  # Include child technologies
    
    # Definition of scenarios and files to check
    scenarios = ['BAU', 'NDCincond', 'NDCcond']
    file_names = ['TotalTechnologyAnnualActivityUpperLimit', 'TotalTechnologyAnnualActivityLowerLimit', 'TotalAnnualMaxCapacity','ResidualCapacity']
    base_path = '1_Baseline_Modelling/'

    # # Iteration over each scenario and file to find missing technologies
    # missing_info = {}
    # for scenario in scenarios:
    #     missing_info[scenario] = {}
    #     for file_name in file_names:
    #         missing = find_missing_technologies(scenario, file_name+'.csv')
    #         missing_info[scenario][file_name] = missing
    
    # missing_info
    
    # ###### TEST 0 ######
    # If i want to know about what mothers o daughters techs are not there      
    # # Calling the function with the provided data
    # check_tech_presence(missing_info, fleet_groups)
    # # Calling the function with the provided data
    # check_mother_absence(missing_info, fleet_groups)


    ###########################################################################################################################
    # Define the path to the text file
    for scen in scenarios:
        # Specify the directory path here
        if specific_case:
            dir_files = os.path.dirname(os.path.abspath(__file__))
        elif tier == str(1):
            dir_files = os.path.join(initial_path, 'Executables')
        else:
            path_of_dir_files = os.path.join(initial_path, 'Futures')
            dir_files = os.path.join(path_of_dir_files, scen)
    
        # List all entries in the directory
        all_entries = os.listdir(dir_files)
    
        # Filter only the directories
        folders = [entry for entry in all_entries if os.path.isdir(os.path.join(dir_files, entry))]
    
        if base_case_RDM or specific_case:
            folders = [folders[0]]
            
        for f in folders:
            if tier == str(1):
                future = 0
            else:
                if base_case_RDM:
                    future = 0
                else:
                    future = f
            
            if future == 0 and not specific_case:
                file_path = os.path.join(initial_path, 'Executables')
                file_path = os.path.join(file_path, f'{scen}_{future}')
                file_path = os.path.join(file_path, f'{scen}_{future}.txt')                
            elif specific_case:
                file_path = specific_case_path_txt
            else:
                file_path = os.path.join(initial_path, 'Futures')
                file_path = os.path.join(file_path, f'{scen}')
                file_path = os.path.join(file_path, f'{future}')
                file_path = os.path.join(file_path, f'{future}.txt')
            
            with open(file_path, 'r') as file:
                lines = file.readlines()
        
            # Define the path to the output file
            if not os.path.exists('tests_results'):
                os.makedirs('tests_results')
            if specific_case:
                output_filename = os.path.join('tests_results', 'comparison_results_{specific_case_path_txt}')
            elif future == 0:
                output_filename = os.path.join(initial_path, 'tests_results')
                output_filename = os.path.join(output_filename, f'comparison_results_{scen}_{future}.txt')
            else:
                output_filename = os.path.join(initial_path, 'tests_results')
                output_filename = os.path.join(output_filename, f'comparison_results_{future}.txt')
            
            
            for i in range(len(file_names)):
                # Take techs defined for the parameter
                result = 0
                result = read_parameters(file_path, file_names[i])  # Make sure parameter is defined
        
                if result is not None:
                    # ###### TEST 1 ######
                    # Check mother - daughters diferrences
                    if file_names[i] == 'ResidualCapacity' or file_names[i] == 'TotalAnnualMaxCapacity':
                        pass
                    else:
                        comparison_results = compare_mother_daughters(file_names[i], result, fleet_groups, output_filename, i, 1)
                    if file_names[i] == 'TotalAnnualMaxCapacity':
                        # ###### TEST 2 ######
                        # Check for decreasing values specifically for this parameter
                        check_any_decreasing_values(result, file_names[i], output_filename, 2)
                else:
                    print("No results available to compare.")
                
                # Only to check the technologies have this parameter
                if file_names[i] == 'ResidualCapacity' and scen == 'LTS' and future == 1:
                    tech_param_check = result
                    
                if file_names[i] == 'TotalTechnologyAnnualActivityUpperLimit':
                    upper_techs = result
                elif file_names[i] == 'TotalTechnologyAnnualActivityLowerLimit':
                    lower_techs = result
                elif file_names[i] == 'TotalAnnualMaxCapacity':
                    maxcapa_techs = result
                elif file_names[i] == 'ResidualCapacity':
                    resicapa_techs = result
                
            # ###### TEST 3 ######
            # Check if values of TotalTechnologyAnnualActivityLowerLimit 
            # is greater than TotalTechnologyAnnualActivityUpperLimit
            compare_techs(upper_techs, lower_techs, output_filename, ['TotalTechnologyAnnualActivityLowerLimit', 'TotalTechnologyAnnualActivityUpperLimit'], 3)
            
            
            # ###### TEST 4 ######
            # Check if values of ResidualCapacity 
            # is greater than TotalAnnualMaxCapacity
            compare_techs(maxcapa_techs, resicapa_techs, output_filename, ['ResidualCapacity', 'TotalAnnualMaxCapacity'], 4)
            
            ###########################################################################################################################
            
            # Load the Excel file
            if tier == str(1):
                path_classifier_modes_transport = os.path.join('A1_Inputs', 'A-I_Classifier_Modes_Transport.xlsx')
                file_path_modes_transp = os.path.join(initial_path, path_classifier_modes_transport)
            elif specific_case:
                file_path_modes_transp = specific_case_file_path_modes_transp
            else:
                path_classifier_modes_transport = os.path.join('0_From_Confection', 'A-I_Classifier_Modes_Transport.xlsx')
                file_path_modes_transp = os.path.join(initial_path, path_classifier_modes_transport)
            
            # Load the 'Mode_Broad' sheet
            df_mode_broad = pd.read_excel(file_path_modes_transp, sheet_name='Mode_Broad')
            
            
            # Take techs defined for the parameter
            ta_maxcap_techs = read_parameters(file_path, 'TotalAnnualMaxCapacity')  # Make sure parameter is defined
            #print(ta_maxcap_techs)
            #sys.exit()
            spec_an_dem_techs = read_parameters(file_path, 'SpecifiedAnnualDemand')  # Make sure parameter is defined
            oar_techs = read_parameters_variant(file_path, 'OutputActivityRatio')  # Make sure parameter is defined
            lower_limit = read_parameters(file_path, 'TotalTechnologyAnnualActivityLowerLimit')  # Make sure parameter is defined
            upper_limit = read_parameters(file_path, 'TotalTechnologyAnnualActivityUpperLimit')  # Make sure parameter is defined
            capfac_techs = read_parameters_variant(file_path, 'CapacityFactor')  # Make sure parameter is defined
            avai_fac_techs = read_parameters(file_path, 'AvailabilityFactor')  # Make sure parameter is defined
            
            # ###### TEST 5 ######
            # Check if SpecifiedAnnualDemand is less than TotalAnnualMaxCapacity x OutputActivityRatio
            check_demand_vs_capacity(df_mode_broad, ta_maxcap_techs, spec_an_dem_techs, oar_techs, output_filename, ['SpecifiedAnnualDemand', 'TotalAnnualMaxCapacity', 'OutputActivityRatio'], 5)
            
            # ###### TEST 6 ######
            # Check if SpecifiedAnnualDemand is less than TotalTechnologyAnnualActivityLowerLimit x OutputActivityRatio
            check_demand_vs_capacity(df_mode_broad, lower_limit, spec_an_dem_techs, oar_techs, output_filename, ['SpecifiedAnnualDemand', 'TotalTechnologyAnnualActivityLowerLimit', 'OutputActivityRatio'], 6)
            
            # ###### TEST 7 ######
            # Check if TotalTechnologyAnnualActivityLowerLimit is less than TotalAnnualMaxCapacity x CapacityFactor x 31.56
            check_demand_vs_capacity_variant(ta_maxcap_techs, capfac_techs, lower_limit, output_filename, ['TotalAnnualMaxCapacity', 'CapacityFactor', 'TotalTechnologyAnnualActivityLowerLimit'], 7)
            
            # ###### TEST 8 ######
            # Check if TotalTechnologyAnnualActivityLowerLimit is less than TotalAnnualMaxCapacity x CapacityFactor x 31.56
            check_demand_vs_capacity_variant(ta_maxcap_techs, capfac_techs, lower_limit, output_filename, ['TotalAnnualMaxCapacity', 'CapacityFactor', 'TotalTechnologyAnnualActivityLowerLimit', 'AvailabilityFactor'], 8, avai_fac_techs)

            # ###### TEST 9 ######
            # Check if TotalAnnualMaxCapacity is less than TotalTechnologyAnnualActivityLowerLimit
            check_maxcapacity_vs_lowerlimit(ta_maxcap_techs, lower_limit, output_filename, ['TotalAnnualMaxCapacity','TotalTechnologyAnnualActivityLowerLimit'], 9)

            # ###### TEST 10 ######
            # Check if TotalAnnualMaxCapacity(daughters) is less than TotalTechnologyAnnualActivityLowerLimit(mother)
            compare_mother_daughters_variant(lower_limit, ta_maxcap_techs, fleet_groups, output_filename, ['TotalAnnualMaxCapacity','TotalTechnologyAnnualActivityLowerLimit'], 10)
            

            # ###### TEST 11 ######
            # Create dict with structure mothers/daughters of the AFOLU sector
            afolu_groups = {}
            for mom in params['all_covers_AFOLU']:
                mom_sector = mom[3:]
                afolu_groups[mom] = params[f'under_covers_{mom[3:]}']
            
            # Check if TotalTechnologyAnnualActivityLowerLimit(daughters) is less than TotalTechnologyAnnualActivityUpperLimit(mother)
            compare_mother_daughters_variant_2(upper_limit, lower_limit, afolu_groups, output_filename, ['TotalTechnologyAnnualActivityLowerLimit','TotalTechnologyAnnualActivityUpperLimit'], 11) 



print('Tests are success.')             