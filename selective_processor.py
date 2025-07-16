import pandas as pd
import numpy as np

# Load full data and test template
full_data_path = 'RK73H_Full_Data.xlsx'
test_template_path = 'test1.xlsx'

full_data = pd.read_excel(full_data_path)
test_template = pd.read_excel(test_template_path)

# Define only the parameters you want to fill
REQUIRED_PARAMETERS = {
    "Resistance": "Resistance",
    "Maximum Working Voltage": "Max Working Voltage (V)",
    "Tolerance": "Tolerance (%)",
    "Package Size": "EIA Code",
    "Rated Power per Element": "Power Rating (W)"
}

def get_selected_part_specs(part_numbers, test_template, full_data, required_params=REQUIRED_PARAMETERS):
    """
    Fill only the specified parameters in the template
    """
    output_frames = []

    for part in part_numbers:
        # Match the part number (exact match first, then contains match)
        row = full_data[full_data['Part Number'] == part]
        if row.empty:
            # Try case-insensitive contains match
            row = full_data[full_data['Part Number'].str.contains(part, case=False, na=False)]
            if row.empty:
                print(f"❌ Part number '{part}' not found.")
                continue

        row = row.iloc[0]  # Get the first match
        print(f"✅ Found data for part number: {part}")

        # Copy template
        filled_template = test_template.copy()
        
        # Clean up the column name (remove trailing space)
        if 'value ' in filled_template.columns:
            filled_template = filled_template.rename(columns={'value ': 'value'})
        
        # Only fill the required parameters
        for idx, param in filled_template["parameter"].items():
            if pd.notna(param) and param in required_params:
                column_name = required_params[param]
                value = row.get(column_name, "")
                filled_template.at[idx, 'value'] = value
                print(f"  📝 {param}: {value}")

        # Add part number as the first column
        filled_template.insert(0, 'Part Number', part)
        output_frames.append(filled_template)

    # Merge all into one DataFrame
    if output_frames:
        final_output = pd.concat(output_frames, ignore_index=True)
        return final_output
    else:
        print("❌ No valid part numbers found.")
        return pd.DataFrame()

def customize_parameters():
    """
    Allow user to select which parameters to fill
    """
    print("\n📋 Available parameters in template:")
    template_params = test_template['parameter'].dropna().tolist()
    
    for i, param in enumerate(template_params, 1):
        print(f"  {i}. {param}")
    
    print("\n📋 Available columns in full data:")
    data_columns = full_data.columns.tolist()
    for i, col in enumerate(data_columns, 1):
        print(f"  {i}. {col}")
    
    print("\nCurrent mapping:")
    for template_param, data_col in REQUIRED_PARAMETERS.items():
        print(f"  '{template_param}' → '{data_col}'")

# Main execution
if __name__ == "__main__":
    print("🔍 Selective Part Number Data Extraction Tool")
    print("="*60)
    
    # Show current parameter selection
    print("📌 Currently filling these parameters:")
    for param in REQUIRED_PARAMETERS.keys():
        print(f"  ✓ {param}")
    
    # Part numbers to process
    part_numbers_list = [
        "RK73H2B TD 1003 FT",
        "RK73H1E TPL 4731 DT"
    ]
    
    print(f"\n🔄 Processing {len(part_numbers_list)} part numbers...")
    
    # Generate the result with only selected parameters
    result_df = get_selected_part_specs(part_numbers_list, test_template, full_data)
    
    if not result_df.empty:
        # Save to Excel
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"selected_specifications_{timestamp}.xlsx"
        
        result_df.to_excel(output_filename, index=False)
        print(f"\n✅ Results saved to '{output_filename}'")
        print(f"📊 Total rows: {len(result_df)}")
        
        # Show summary
        print(f"\n📋 Filled parameters for each part:")
        for param in REQUIRED_PARAMETERS.keys():
            count = result_df[result_df['parameter'] == param]['value'].notna().sum()
            print(f"  {param}: {count} values filled")
    else:
        print("❌ No data to save.")
