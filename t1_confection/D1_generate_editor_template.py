"""
Generate Secondary Techs Editor Template

This script reads all A-O_Parametrization.xlsx files and generates
a user-friendly Excel template for editing Secondary Techs data.

Usage:
    python t1_confection/generate_editor_template.py
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Protection
from openpyxl.worksheet.datavalidation import DataValidation
import sys
from pathlib import Path
from datetime import datetime


def collect_data_from_all_scenarios():
    """
    Collect all unique values from all scenario files

    Returns:
        dict with scenarios, countries, tech_mapping (Tech.Name -> Tech), parameters, years
    """
    base_path = Path(__file__).parent / "A1_Outputs"

    scenarios = ["BAU", "NDC", "NDC+ELC", "NDC_NoRPO"]

    all_countries = set()
    tech_mapping = {}  # Tech.Name -> Tech code
    all_parameters = set()
    all_years = set()

    print("Collecting data from all scenarios...")
    print()

    for scenario in scenarios:
        scenario_path = base_path / f"A1_Outputs_{scenario}" / "A-O_Parametrization.xlsx"

        if not scenario_path.exists():
            print(f"WARNING: File not found: {scenario_path}")
            continue

        print(f"Reading {scenario}...")

        wb = openpyxl.load_workbook(scenario_path, data_only=True)

        if 'Secondary Techs' not in wb.sheetnames:
            print(f"  WARNING: 'Secondary Techs' sheet not found in {scenario}")
            wb.close()
            continue

        ws = wb['Secondary Techs']

        # Get headers to find year columns
        headers = [cell.value for cell in ws[1]]

        # Find year columns (columns with numeric values representing years)
        year_col_indices = []
        for idx, header in enumerate(headers, 1):
            if header and str(header).isdigit():
                try:
                    year = int(header)
                    if 2000 <= year <= 2100:
                        all_years.add(year)
                        year_col_indices.append(idx)
                except:
                    pass

        # Collect tech mapping and parameters (skip header row)
        # Columns: 1=Tech.Id, 2=Tech, 3=Tech.Name, 5=Parameter
        for row_idx in range(2, ws.max_row + 1):
            tech_code = ws.cell(row_idx, 2).value      # Column 2: Tech
            tech_name = ws.cell(row_idx, 3).value      # Column 3: Tech.Name
            parameter = ws.cell(row_idx, 5).value      # Column 5: Parameter

            if tech_code and tech_name:
                tech_code_str = str(tech_code).strip()
                tech_name_str = str(tech_name).strip()

                # Build mapping: Tech.Name -> Tech
                tech_mapping[tech_name_str] = tech_code_str

                # Extract country code from PWR technologies only
                # Format: PWRTRNARGXX -> country code is ARG (characters 6-8)
                if tech_code_str.upper().startswith('PWR') and len(tech_code_str) >= 9:
                    country_code = tech_code_str[6:9].upper()
                    all_countries.add(country_code)

            if parameter:
                all_parameters.add(str(parameter).strip())

        wb.close()
        print(f"  Found: {len(tech_mapping)} technologies, {len(all_parameters)} parameters")

    print()
    print(f"Summary:")
    print(f"  Scenarios: {len(scenarios)}")
    print(f"  Countries: {len(all_countries)} - {sorted(all_countries)}")
    print(f"  Technologies (Tech.Name): {len(tech_mapping)}")
    print(f"  Parameters: {len(all_parameters)}")
    print(f"  Years: {len(all_years)} - {min(all_years) if all_years else 'N/A'} to {max(all_years) if all_years else 'N/A'}")
    print()

    return {
        'scenarios': sorted(scenarios),
        'countries': sorted(all_countries),
        'tech_mapping': tech_mapping,  # Tech.Name -> Tech code
        'tech_names': sorted(tech_mapping.keys()),  # List of Tech.Name for dropdown
        'parameters': sorted(all_parameters),
        'years': sorted(all_years)
    }


def create_editor_template(data, output_path):
    """
    Create the Excel editor template with dropdowns and validation

    Args:
        data: dict with scenarios, countries, technologies, parameters, years
        output_path: Path where to save the template
    """
    print("Creating Excel template...")

    wb = openpyxl.Workbook()

    # Main sheet for data entry
    ws_main = wb.active
    ws_main.title = "Editor"

    # Create hidden sheets for dropdown lists and mappings
    ws_scenarios = wb.create_sheet("_Scenarios")
    ws_countries = wb.create_sheet("_Countries")
    ws_tech_names = wb.create_sheet("_TechNames")
    ws_tech_mapping = wb.create_sheet("_TechMapping")  # Tech.Name -> Tech code
    ws_parameters = wb.create_sheet("_Parameters")

    # Hide validation sheets
    ws_scenarios.sheet_state = 'hidden'
    ws_countries.sheet_state = 'hidden'
    ws_tech_names.sheet_state = 'hidden'
    ws_tech_mapping.sheet_state = 'hidden'
    ws_parameters.sheet_state = 'hidden'

    # Populate validation sheets
    for idx, scenario in enumerate(['ALL'] + data['scenarios'], 1):
        ws_scenarios.cell(idx, 1, scenario)

    for idx, country in enumerate(data['countries'], 1):
        ws_countries.cell(idx, 1, country)

    # Populate Tech.Name list and Tech.Name -> Tech mapping
    for idx, (tech_name, tech_code) in enumerate(sorted(data['tech_mapping'].items()), 1):
        ws_tech_names.cell(idx, 1, tech_name)  # Column A: Tech.Name for dropdown
        ws_tech_mapping.cell(idx, 1, tech_name)  # Column A: Tech.Name
        ws_tech_mapping.cell(idx, 2, tech_code)  # Column B: Tech code

    for idx, param in enumerate(data['parameters'], 1):
        ws_parameters.cell(idx, 1, param)

    # Create header row in main sheet
    # Columns: Scenario, Country, Tech.Name, Tech (auto-filled), Parameter, Years...
    headers = ['Scenario', 'Country', 'Tech.Name', 'Tech', 'Parameter'] + [str(year) for year in data['years']]

    # Style for header
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, header in enumerate(headers, 1):
        cell = ws_main.cell(1, col_idx, header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    # Set column widths
    ws_main.column_dimensions['A'].width = 15  # Scenario
    ws_main.column_dimensions['B'].width = 12  # Country
    ws_main.column_dimensions['C'].width = 40  # Tech.Name
    ws_main.column_dimensions['D'].width = 20  # Tech (auto-filled, read-only)
    ws_main.column_dimensions['E'].width = 30  # Parameter

    # Year columns (start from column F = 6)
    for col_idx in range(6, len(headers) + 1):
        ws_main.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 10

    # Add data validations (dropdowns) for first 100 rows
    max_data_rows = 100

    # Scenario dropdown (column A)
    dv_scenario = DataValidation(
        type="list",
        formula1=f"=_Scenarios!$A$1:$A${len(data['scenarios']) + 1}",
        allow_blank=False
    )
    dv_scenario.error = 'Please select a valid scenario'
    dv_scenario.errorTitle = 'Invalid Scenario'
    ws_main.add_data_validation(dv_scenario)
    dv_scenario.add(f'A2:A{max_data_rows + 1}')

    # Country dropdown (column B)
    dv_country = DataValidation(
        type="list",
        formula1=f"=_Countries!$A$1:$A${len(data['countries'])}",
        allow_blank=False
    )
    dv_country.error = 'Please select a valid country'
    dv_country.errorTitle = 'Invalid Country'
    ws_main.add_data_validation(dv_country)
    dv_country.add(f'B2:B{max_data_rows + 1}')

    # Tech.Name dropdown (column C)
    dv_tech_name = DataValidation(
        type="list",
        formula1=f"=_TechNames!$A$1:$A${len(data['tech_names'])}",
        allow_blank=False
    )
    dv_tech_name.error = 'Please select a valid technology name'
    dv_tech_name.errorTitle = 'Invalid Tech.Name'
    ws_main.add_data_validation(dv_tech_name)
    dv_tech_name.add(f'C2:C{max_data_rows + 1}')

    # Column D (Tech) will be auto-filled with VLOOKUP formula
    # Add formulas to map Tech.Name to Tech code
    for row_idx in range(2, max_data_rows + 2):
        formula = f'=IFERROR(VLOOKUP(C{row_idx},_TechMapping!$A:$B,2,FALSE),"")'
        ws_main.cell(row_idx, 4, formula)  # Column D

    # Protect column D (Tech) from manual editing
    for row_idx in range(2, max_data_rows + 2):
        ws_main.cell(row_idx, 4).protection = Protection(locked=True)

    # Parameter dropdown (column E)
    dv_parameter = DataValidation(
        type="list",
        formula1=f"=_Parameters!$A$1:$A${len(data['parameters'])}",
        allow_blank=False
    )
    dv_parameter.error = 'Please select a valid parameter'
    dv_parameter.errorTitle = 'Invalid Parameter'
    ws_main.add_data_validation(dv_parameter)
    dv_parameter.add(f'E2:E{max_data_rows + 1}')

    # Freeze panes (freeze first row and first 5 columns)
    ws_main.freeze_panes = 'F2'

    # Add instructions in a separate sheet
    ws_instructions = wb.create_sheet("Instructions", 0)
    ws_instructions.column_dimensions['A'].width = 80

    instructions = [
        ["Secondary Techs Editor - Instructions", ""],
        ["", ""],
        ["This Excel file allows you to edit Secondary Techs data easily.", ""],
        ["", ""],
        ["HOW TO USE:", ""],
        ["1. Go to the 'Editor' sheet", ""],
        ["2. Fill in each row with:", ""],
        ["   - Scenario: Select BAU, NDC, NDC+ELC, NDC_NoRPO, or ALL (applies to all scenarios)", ""],
        ["   - Country: Select the country code (ARG, BOL, CHI, COL, ECU, GUA, etc.)", ""],
        ["   - Tech.Name: Select the descriptive technology name from the dropdown", ""],
        ["   - Tech: This column is auto-filled based on Tech.Name (READ-ONLY)", ""],
        ["   - Parameter: Select the parameter to modify", ""],
        ["   - Year values: Enter numeric values for each year (leave empty to keep current value)", ""],
        ["", ""],
        ["3. Save and close this file", ""],
        ["4. Run: python t1_confection/update_secondary_techs.py", ""],
        ["", ""],
        ["IMPORTANT NOTES:", ""],
        ["- You can add as many rows as needed", ""],
        ["- The 'Tech' column is automatically filled when you select a 'Tech.Name'", ""],
        ["- Empty year cells will NOT modify the current value in the destination file", ""],
        ["- Scenario 'ALL' will apply changes to all 4 scenario files", ""],
        ["- A backup will be created automatically before making changes", ""],
        ["- Check the log file after running the update script", ""],
        ["", ""],
        [f"Template generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ""],
    ]

    for row_idx, row_data in enumerate(instructions, 1):
        cell = ws_instructions.cell(row_idx, 1, row_data[0])
        if row_idx == 1:
            cell.font = Font(size=16, bold=True, color="366092")
        elif row_data[0].startswith("HOW TO USE:") or row_data[0].startswith("IMPORTANT NOTES:"):
            cell.font = Font(size=12, bold=True)

        cell.alignment = Alignment(wrap_text=True, vertical="top")

    # Save the workbook
    wb.save(output_path)
    print(f"Template saved: {output_path}")
    print()


def main():
    try:
        print("=" * 80)
        print("SECONDARY TECHS EDITOR - TEMPLATE GENERATOR")
        print("=" * 80)
        print()

        # Collect data from all scenarios
        data = collect_data_from_all_scenarios()

        if not data['tech_mapping'] or not data['parameters'] or not data['years']:
            print("ERROR: Could not collect enough data from scenario files")
            print("Please check that A-O_Parametrization.xlsx files exist and contain Secondary Techs sheet")
            return 1

        # Create template
        output_path = Path(__file__).parent / "Secondary_Techs_Editor.xlsx"
        create_editor_template(data, output_path)

        print("=" * 80)
        print("TEMPLATE GENERATION COMPLETE")
        print("=" * 80)
        print()
        print("Next steps:")
        print("1. Open t1_confection/Secondary_Techs_Editor.xlsx")
        print("2. Fill in the 'Editor' sheet with your changes")
        print("3. Save and close the file")
        print("4. Run: python t1_confection/update_secondary_techs.py")
        print()

        return 0

    except Exception as e:
        print(f"\nERROR: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
