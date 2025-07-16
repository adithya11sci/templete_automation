import pandas as pd

def fill_specifications_from_part_numbers(part_numbers):
    """
    Main function to fill test1.xlsx template with data from RK73H_Full_Data.xlsx
    
    Args:
        part_numbers (list): List of part numbers to process
    
    Returns:
        str: Filename of the created Excel file
    """
    
    # Load data files
    print("📂 Loading data files...")
    full_data = pd.read_excel('RK73H_Full_Data.xlsx')
    template = pd.read_excel('test1.xlsx')
    
    # Clean column names
    template.columns = ['parameter', 'unit', 'value']
    
    all_results = []
    
    for i, part_number in enumerate(part_numbers, 1):
        print(f"[{i}/{len(part_numbers)}] Processing: {part_number}")
        
        # Find part in database
        part_row = full_data[full_data['Part Number'] == part_number]
        
        if part_row.empty:
            print(f"   ❌ Not found: {part_number}")
            continue
            
        part_data = part_row.iloc[0]
        print(f"   ✅ Found: {part_number}")
        
        # Create filled template
        filled = template.copy()
        
        # Map data to template
        data_mapping = {
            'Specifications': part_number,
            'Resistance': part_data.get('Resistance', ''),
            'Maximum Working Voltage': part_data.get('Max Working Voltage (V)', ''),
            'Tolerance': part_data.get('Tolerance (%)', ''),
            'Operating Temperature': '-55°C to +155°C',
            'Package Size': part_data.get('EIA Code', ''),
            'Rated Power per Element': part_data.get('Power Rating (W)', ''),
            'Temperature Coefficient': part_data.get('T.C.R. (ppm/°C)', ''),
            'Lead Finish': part_data.get('Termination Material', ''),
            'Technology': 'Thick Film'
        }
        
        # Fill values
        for idx, param in filled['parameter'].items():
            if pd.notna(param) and param in data_mapping:
                filled.at[idx, 'value'] = data_mapping[param]
        
        # Add separator between parts
        if len(all_results) > 0:
            separator = pd.DataFrame({
                'parameter': ['', '--- Next Part ---', ''],
                'unit': ['', '', ''],
                'value': ['', '', '']
            })
            all_results.append(separator)
        
        all_results.append(filled)
    
    # Combine all results
    if all_results:
        final_result = pd.concat(all_results, ignore_index=True)
        
        # Save to Excel
        output_file = 'filled_specifications.xlsx'
        final_result.to_excel(output_file, index=False)
        
        print(f"\n✅ Saved to: {output_file}")
        print(f"📊 Total rows: {len(final_result)}")
        
        return output_file
    else:
        print("❌ No valid parts found")
        return None

# Example usage:
if __name__ == "__main__":
    # Enter your part numbers here
    my_part_numbers = [
        "RK73H2B TD 1003 FT",
        "RK73H1E TPL 4731 DT"
    ]
    
    # Process the part numbers
    result_file = fill_specifications_from_part_numbers(my_part_numbers)
    
    if result_file:
        print(f"\n🎉 Success! Check the file: {result_file}")
    else:
        print("\n❌ Failed to process any parts")
