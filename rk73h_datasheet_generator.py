import pandas as pd
import numpy as np
from datetime import datetime
import re

class RK73HDataProvider:
    """
    Data provider class that uses extracted PDF data to fill templates
    """
    
    def __init__(self):
        self.extracted_data = self.load_extracted_data()
        self.template = self.load_template()
        
    def load_extracted_data(self):
        """Load the extracted data from the PDF"""
        print("📂 Loading extracted RK73H data...")
        
        # This data is based on what was extracted from the PDF
        extracted_specs = {
            # General specifications from PDF
            'series': 'RK73H',
            'type': 'Thick Film Chip Resistor',
            'manufacturer': 'KOA Speer',
            'technology': 'Thick Film',
            
            # Electrical specifications
            'resistance_range': '1Ω to 10MΩ',
            'power_ratings': {
                '0402': '0.063W (1/16W)',
                '0603': '0.1W (1/10W)', 
                '0805': '0.125W (1/8W)',
                '1206': '0.25W (1/4W)',
                '1210': '0.5W (1/2W)',
                '1812': '0.75W (3/4W)',
                '2010': '0.75W (3/4W)',
                '2512': '1W'
            },
            'tolerance_options': ['±0.1%', '±0.25%', '±0.5%', '±1%', '±5%'],
            'voltage_rating': '50V to 200V (depending on size)',
            
            # Temperature specifications
            'operating_temp': '-55°C to +155°C',
            'tcr': '±100/±200/±400 ppm/°C',
            
            # Physical specifications
            'package_sizes': {
                '1E': '0402 (1.0×0.5mm)',
                '1J': '0603 (1.6×0.8mm)',
                '2A': '0805 (2.0×1.25mm)',
                '2B': '1206 (3.2×1.6mm)',
                '2F': '1210 (3.2×2.5mm)',
                '3A': '1812 (4.5×3.2mm)',
                '3B': '2010 (5.0×2.5mm)',
                '3C': '2512 (6.4×3.2mm)'
            },
            
            # Termination
            'termination': 'Cu/Ni/Sn',
            'packaging': 'Tape & Reel',
            
            # Environmental
            'automotive_qualified': 'AEC-Q200',
            'halogen_free': 'Yes',
            
            # Resistance codes (sample)
            'resistance_codes': {
                '1001': '1kΩ',
                '1002': '10kΩ', 
                '1003': '100kΩ',
                '4731': '4.73kΩ',
                '1000': '100Ω',
                '1500': '150Ω'
            }
        }
        
        return extracted_specs
    
    def load_template(self):
        """Load the template structure"""
        try:
            template_df = pd.read_excel('test1.xlsx')
            return template_df
        except:
            # Create a default template if test1.xlsx is not available
            return self.create_default_template()
    
    def create_default_template(self):
        """Create a default template structure"""
        template_data = {
            'parameter': [
                'Specifications',
                'Resistance',
                'Maximum Working Voltage', 
                'Tolerance',
                'Operating Temperature',
                'Package Size',
                'Rated Power per Element',
                'Temperature Coefficient',
                'Lead Finish',
                'Technology',
                'Series',
                'Automotive Qualified',
                'Environmental Compliance',
                'Packaging Type'
            ],
            'unit': [
                '', '[Ohm]', '[V dc]', '[± %]', '[℃]', '[EIA]', '[W]', 
                '[± ppm/K]', '', '', '', '', '', ''
            ],
            'value': ['' for _ in range(14)]
        }
        return pd.DataFrame(template_data)

def decode_part_number(part_number):
    """
    Decode RK73H part number to extract specifications
    """
    print(f"🔍 Decoding part number: {part_number}")
    
    decoded = {
        'series': 'RK73H',
        'size_code': '',
        'resistance_code': '',
        'tolerance_code': '',
        'termination_code': '',
        'packaging_code': ''
    }
    
    # Clean the part number
    clean_part = re.sub(r'\s+', '', part_number.upper())
    
    if clean_part.startswith('RK73H'):
        # Extract size code (positions 6-7)
        if len(clean_part) > 7:
            decoded['size_code'] = clean_part[5:7]
        
        # Extract other codes based on typical RK73H structure
        # This is based on the PDF structure analysis
        parts = clean_part.split()
        if len(parts) >= 4:
            decoded['resistance_code'] = parts[2] if len(parts) > 2 else ''
            decoded['tolerance_code'] = parts[3][0] if len(parts) > 3 else ''
    
    return decoded

def fill_template_with_part_data(part_number, data_provider):
    """
    Fill template with data for a specific part number
    """
    print(f"📝 Filling template for: {part_number}")
    
    # Get template
    template = data_provider.template.copy()
    extracted_data = data_provider.extracted_data
    
    # Decode part number
    decoded = decode_part_number(part_number)
    
    # Get size-specific power rating
    size_code = decoded['size_code']
    power_rating = ''
    package_size = ''
    
    # Map size codes to EIA codes and power ratings
    size_mapping = {
        '1E': ('0402', '0.063W'),
        '1J': ('0603', '0.1W'),
        '2A': ('0805', '0.125W'),
        '2B': ('1206', '0.25W'),
        '2F': ('1210', '0.5W'),
        '3A': ('1812', '0.75W'),
        '3B': ('2010', '0.75W'),
        '3C': ('2512', '1W')
    }
    
    if size_code in size_mapping:
        package_size, power_rating = size_mapping[size_code]
    
    # Get resistance value
    resistance_code = decoded['resistance_code']
    resistance_value = extracted_data['resistance_codes'].get(resistance_code, 'See datasheet')
    
    # Fill the template
    parameter_values = {
        'Specifications': part_number,
        'Resistance': resistance_value,
        'Maximum Working Voltage': '50V',  # Default, varies by size
        'Tolerance': '±1%',  # Default
        'Operating Temperature': extracted_data['operating_temp'],
        'Package Size': package_size,
        'Rated Power per Element': power_rating,
        'Temperature Coefficient': '±200 ppm/°C',  # Default
        'Lead Finish': extracted_data['termination'],
        'Technology': extracted_data['technology'],
        'Series': extracted_data['series'],
        'Automotive Qualified': extracted_data['automotive_qualified'],
        'Environmental Compliance': 'Halogen-Free',
        'Packaging Type': extracted_data['packaging']
    }
    
    # Apply values to template
    for idx, param in template['parameter'].items():
        if pd.notna(param) and param in parameter_values:
            template.at[idx, 'value'] = parameter_values[param]
    
    return template

def process_multiple_parts(part_numbers_list):
    """
    Process multiple part numbers and create filled templates
    """
    print("🚀 Starting batch processing...")
    print("=" * 50)
    
    # Initialize data provider
    data_provider = RK73HDataProvider()
    
    all_results = []
    
    for i, part_number in enumerate(part_numbers_list, 1):
        print(f"\n[{i}/{len(part_numbers_list)}] Processing: {part_number}")
        
        # Fill template for this part
        filled_template = fill_template_with_part_data(part_number, data_provider)
        
        # Add part number identifier
        filled_template.insert(0, 'Part_Number', part_number)
        
        # Add separator between parts
        if len(all_results) > 0:
            separator = pd.DataFrame({
                'Part_Number': [''],
                'parameter': ['--- Next Part ---'],
                'unit': [''],
                'value': ['']
            })
            all_results.append(separator)
        
        all_results.append(filled_template)
        print(f"      ✅ Template filled successfully")
    
    # Combine all results
    if all_results:
        final_result = pd.concat(all_results, ignore_index=True)
        return final_result
    else:
        print("❌ No results to combine")
        return pd.DataFrame()

def save_filled_datasheet(result_df, filename=None):
    """
    Save the filled datasheet to Excel
    """
    if filename is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"RK73H_Filled_Datasheet_{timestamp}.xlsx"
    
    print(f"\n💾 Saving filled datasheet: {filename}")
    
    result_df.to_excel(filename, index=False, sheet_name='Filled_Specifications')
    
    print(f"✅ Datasheet saved: {filename}")
    print(f"📊 Total rows: {len(result_df)}")
    
    return filename

def get_user_input():
    """
    Get part numbers from user input
    """
    print("\n📝 Enter RK73H part numbers to process:")
    print("Examples:")
    print("  - RK73H2B TD 1003 FT")
    print("  - RK73H1E TPL 4731 DT")
    print("\nType each part number and press Enter. Type 'done' when finished.\n")
    
    part_numbers = []
    while True:
        part_input = input("Part number: ").strip()
        if part_input.lower() == 'done':
            break
        if part_input:
            part_numbers.append(part_input)
            print(f"  ✓ Added: {part_input}")
    
    return part_numbers

# Main execution function
def main():
    """
    Main function to handle user input and generate filled datasheet
    """
    print("🔧 RK73H Datasheet Generator")
    print("Using extracted PDF data to fill templates")
    print("=" * 60)
    
    # Option 1: Use example part numbers
    use_examples = input("\nUse example part numbers? (y/n): ").lower().startswith('y')
    
    if use_examples:
        part_numbers = [
            "RK73H2B TD 1003 FT",
            "RK73H1E TPL 4731 DT",
            "RK73H1J TK 1001 FT",
            "RK73H2A TP 1002 DT"
        ]
        print(f"Using example parts: {part_numbers}")
    else:
        # Option 2: Get user input
        part_numbers = get_user_input()
    
    if not part_numbers:
        print("❌ No part numbers provided. Exiting.")
        return
    
    # Process the part numbers
    print(f"\n🔄 Processing {len(part_numbers)} part numbers...")
    result = process_multiple_parts(part_numbers)
    
    if not result.empty:
        # Save the filled datasheet
        output_file = save_filled_datasheet(result)
        
        print(f"\n🎉 SUCCESS!")
        print(f"📁 Filled datasheet created: {output_file}")
        print(f"📋 Based on extracted RK73H PDF data")
        
        return output_file
    else:
        print("❌ Failed to create filled datasheet")
        return None

if __name__ == "__main__":
    main()
