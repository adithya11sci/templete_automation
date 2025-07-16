import pandas as pd
import pdfplumber
import PyPDF2
import re
from io import StringIO

def extract_pdf_data(pdf_path):
    """Extract all data from RK73H.pdf and organize it"""
    
    print("📖 Reading PDF file...")
    
    # Initialize data storage
    extracted_data = {
        'specifications': [],
        'part_numbers': [],
        'electrical_characteristics': [],
        'physical_dimensions': [],
        'ordering_information': []
    }
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_text = ""
            tables = []
            
            print(f"📄 PDF has {len(pdf.pages)} pages")
            
            # Extract text and tables from each page
            for page_num, page in enumerate(pdf.pages, 1):
                print(f"📝 Processing page {page_num}...")
                
                # Extract text
                page_text = page.extract_text()
                if page_text:
                    all_text += f"\n--- PAGE {page_num} ---\n" + page_text
                
                # Extract tables
                page_tables = page.extract_tables()
                if page_tables:
                    for table_num, table in enumerate(page_tables):
                        print(f"  📊 Found table {table_num + 1} on page {page_num}")
                        tables.append({
                            'page': page_num,
                            'table_num': table_num + 1,
                            'data': table
                        })
            
            # Process extracted data
            print("\n🔍 Analyzing extracted content...")
            
            # Extract specifications
            specs = extract_specifications(all_text)
            extracted_data['specifications'] = specs
            
            # Extract part numbers and their details
            part_info = extract_part_numbers(all_text)
            extracted_data['part_numbers'] = part_info
            
            # Process tables for structured data
            table_data = process_tables(tables)
            extracted_data.update(table_data)
            
            # Extract electrical characteristics
            electrical = extract_electrical_characteristics(all_text)
            extracted_data['electrical_characteristics'] = electrical
            
            # Extract physical dimensions
            dimensions = extract_dimensions(all_text)
            extracted_data['physical_dimensions'] = dimensions
            
            return extracted_data, all_text
            
    except Exception as e:
        print(f"❌ Error reading PDF: {e}")
        return None, None

def extract_specifications(text):
    """Extract general specifications from text"""
    specs = []
    
    # Common specification patterns
    spec_patterns = [
        r'Resistance[:\s]+([^\\n]+)',
        r'Tolerance[:\s]+([^\\n]+)',
        r'Power Rating[:\s]+([^\\n]+)',
        r'Working Voltage[:\s]+([^\\n]+)',
        r'Temperature Range[:\s]+([^\\n]+)',
        r'Temperature Coefficient[:\s]+([^\\n]+)',
        r'Package[:\s]+([^\\n]+)',
        r'Series[:\s]+([^\\n]+)'
    ]
    
    for pattern in spec_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        if matches:
            param_name = pattern.split('[')[0]
            for match in matches:
                specs.append({
                    'Parameter': param_name,
                    'Value': match.strip()
                })
    
    return specs

def extract_part_numbers(text):
    """Extract part numbers and their details"""
    part_numbers = []
    
    # Pattern for RK73H part numbers
    part_pattern = r'RK73H[A-Z0-9\s]+[A-Z]{2,3}\s*[0-9]{3,4}\s*[A-Z]{1,2}'
    
    matches = re.findall(part_pattern, text)
    for match in matches:
        part_numbers.append({
            'Part Number': match.strip(),
            'Series': 'RK73H',
            'Type': 'Thick Film Resistor'
        })
    
    return part_numbers

def extract_electrical_characteristics(text):
    """Extract electrical characteristics"""
    characteristics = []
    
    # Look for electrical parameter sections
    electrical_patterns = [
        r'Resistance Range[:\s]+([^\\n]+)',
        r'Power Dissipation[:\s]+([^\\n]+)',
        r'Voltage Rating[:\s]+([^\\n]+)',
        r'Temperature Coefficient[:\s]+([^\\n]+)',
        r'Tolerance[:\s]+([^\\n]+)'
    ]
    
    for pattern in electrical_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        if matches:
            param_name = pattern.split('[')[0]
            for match in matches:
                characteristics.append({
                    'Parameter': param_name,
                    'Specification': match.strip()
                })
    
    return characteristics

def extract_dimensions(text):
    """Extract physical dimensions"""
    dimensions = []
    
    # Look for dimension patterns
    dim_patterns = [
        r'Length[:\s]+([0-9.]+\s*mm)',
        r'Width[:\s]+([0-9.]+\s*mm)',
        r'Height[:\s]+([0-9.]+\s*mm)',
        r'Thickness[:\s]+([0-9.]+\s*mm)',
        r'([0-9.]+)\s*×\s*([0-9.]+)\s*×?\s*([0-9.]*)\s*mm'
    ]
    
    for pattern in dim_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        if matches:
            for match in matches:
                if isinstance(match, tuple):
                    dimensions.append({
                        'Dimension Type': 'L×W×H',
                        'Value': ' × '.join([str(m) for m in match if m])
                    })
                else:
                    dimensions.append({
                        'Dimension Type': 'Physical',
                        'Value': match.strip()
                    })
    
    return dimensions

def process_tables(tables):
    """Process extracted tables into structured data"""
    processed_tables = {}
    
    for table_info in tables:
        table_data = table_info['data']
        page = table_info['page']
        table_num = table_info['table_num']
        
        if table_data and len(table_data) > 1:
            # Try to create a DataFrame from the table
            try:
                df = pd.DataFrame(table_data[1:], columns=table_data[0])
                processed_tables[f'table_page_{page}_{table_num}'] = df.to_dict('records')
            except:
                # If that fails, store as raw data
                processed_tables[f'raw_table_page_{page}_{table_num}'] = table_data
    
    return processed_tables

def create_excel_datasheet(extracted_data, output_file='RK73H_Complete_Datasheet.xlsx'):
    """Create comprehensive Excel datasheet"""
    
    print(f"\n📊 Creating Excel datasheet: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        
        # Sheet 1: Specifications
        if extracted_data['specifications']:
            specs_df = pd.DataFrame(extracted_data['specifications'])
            specs_df.to_excel(writer, sheet_name='Specifications', index=False)
            print("  ✓ Specifications sheet created")
        
        # Sheet 2: Part Numbers
        if extracted_data['part_numbers']:
            parts_df = pd.DataFrame(extracted_data['part_numbers'])
            parts_df.to_excel(writer, sheet_name='Part_Numbers', index=False)
            print("  ✓ Part Numbers sheet created")
        
        # Sheet 3: Electrical Characteristics
        if extracted_data['electrical_characteristics']:
            electrical_df = pd.DataFrame(extracted_data['electrical_characteristics'])
            electrical_df.to_excel(writer, sheet_name='Electrical_Characteristics', index=False)
            print("  ✓ Electrical Characteristics sheet created")
        
        # Sheet 4: Physical Dimensions
        if extracted_data['physical_dimensions']:
            dimensions_df = pd.DataFrame(extracted_data['physical_dimensions'])
            dimensions_df.to_excel(writer, sheet_name='Physical_Dimensions', index=False)
            print("  ✓ Physical Dimensions sheet created")
        
        # Additional sheets for tables
        sheet_count = 5
        for key, data in extracted_data.items():
            if key.startswith('table_') and data:
                try:
                    table_df = pd.DataFrame(data)
                    sheet_name = f'Table_{sheet_count}'
                    table_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"  ✓ {sheet_name} created")
                    sheet_count += 1
                except:
                    pass
    
    print(f"\n✅ Excel datasheet created: {output_file}")
    return output_file

# Main execution
if __name__ == "__main__":
    print("📋 RK73H PDF Data Extraction")
    print("=" * 50)
    
    pdf_file = 'RK73H.pdf'
    
    # Extract data from PDF
    extracted_data, full_text = extract_pdf_data(pdf_file)
    
    if extracted_data:
        # Create Excel datasheet
        excel_file = create_excel_datasheet(extracted_data)
        
        # Save full text for reference
        with open('RK73H_extracted_text.txt', 'w', encoding='utf-8') as f:
            f.write(full_text)
        
        print("\n📁 Files created:")
        print(f"  📊 Excel datasheet: {excel_file}")
        print(f"  📝 Full text: RK73H_extracted_text.txt")
        
        # Show summary
        print("\n📋 Extraction Summary:")
        for key, data in extracted_data.items():
            if isinstance(data, list):
                print(f"  {key}: {len(data)} items")
            elif isinstance(data, dict):
                print(f"  {key}: {len(data)} entries")
    else:
        print("❌ Failed to extract data from PDF")
