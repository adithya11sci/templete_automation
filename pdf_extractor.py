import pandas as pd
import pdfplumber
import re
from datetime import datetime
import numpy as np

def extract_pdf_data(pdf_path):
    """
    Extract all information from RK73H.pdf and structure it into organized data
    """
    print("🔍 Extracting data from PDF...")
    
    extracted_data = {
        'text_content': [],
        'tables': [],
        'specifications': {},
        'part_numbers': [],
        'technical_data': []
    }
    
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        print(f"📄 Processing {total_pages} pages...")
        
        for page_num, page in enumerate(pdf.pages, 1):
            print(f"   Processing page {page_num}/{total_pages}")
            
            # Extract text
            text = page.extract_text()
            if text:
                extracted_data['text_content'].append({
                    'page': page_num,
                    'text': text
                })
            
            # Extract tables
            tables = page.extract_tables()
            for table_idx, table in enumerate(tables):
                if table:
                    extracted_data['tables'].append({
                        'page': page_num,
                        'table_index': table_idx,
                        'data': table
                    })
    
    return extracted_data

def parse_specifications(extracted_data):
    """
    Parse specifications from the extracted text
    """
    print("📋 Parsing specifications...")
    
    specifications = {}
    all_text = ' '.join([item['text'] for item in extracted_data['text_content']])
    
    # Common specification patterns
    spec_patterns = {
        'resistance_range': r'Resistance.*?(\d+.*?Ω.*?\d+.*?Ω)',
        'tolerance': r'Tolerance.*?([±]?\d+\.?\d*%)',
        'power_rating': r'Power.*?(\d+\.?\d*\s*W)',
        'voltage': r'Voltage.*?(\d+\.?\d*\s*V)',
        'temperature_range': r'Temperature.*?(-?\d+°C.*?\+?\d+°C)',
        'tcr': r'T\.C\.R.*?([±]?\d+.*?ppm)',
        'package_sizes': r'Package.*?(EIA.*?\d+)',
        'series': r'Series.*?(RK\d+[A-Z]*)'
    }
    
    for spec_name, pattern in spec_patterns.items():
        matches = re.findall(pattern, all_text, re.IGNORECASE)
        if matches:
            specifications[spec_name] = matches
    
    return specifications

def extract_part_numbers(extracted_data):
    """
    Extract all part numbers from the PDF
    """
    print("🔢 Extracting part numbers...")
    
    part_numbers = []
    all_text = ' '.join([item['text'] for item in extracted_data['text_content']])
    
    # RK73H part number pattern
    part_pattern = r'RK73H[0-9A-Z\s]{10,20}[A-Z]{1,3}'
    matches = re.findall(part_pattern, all_text)
    
    for match in matches:
        clean_part = re.sub(r'\s+', ' ', match.strip())
        if clean_part not in part_numbers:
            part_numbers.append(clean_part)
    
    return part_numbers

def process_tables(extracted_data):
    """
    Process and structure table data
    """
    print("📊 Processing tables...")
    
    processed_tables = []
    
    for table_info in extracted_data['tables']:
        table_data = table_info['data']
        page = table_info['page']
        
        if not table_data or len(table_data) < 2:
            continue
        
        # Try to identify table structure
        headers = table_data[0]
        rows = table_data[1:]
        
        # Clean headers
        clean_headers = []
        for header in headers:
            if header:
                clean_headers.append(str(header).strip())
            else:
                clean_headers.append(f"Column_{len(clean_headers)}")
        
        # Create DataFrame
        try:
            df = pd.DataFrame(rows, columns=clean_headers)
            
            # Remove empty rows
            df = df.dropna(how='all')
            
            # Store processed table
            processed_tables.append({
                'page': page,
                'table_index': table_info['table_index'],
                'dataframe': df,
                'headers': clean_headers,
                'row_count': len(df)
            })
            
        except Exception as e:
            print(f"   ⚠️ Error processing table on page {page}: {e}")
            continue
    
    return processed_tables

def create_comprehensive_datasheet(extracted_data, specifications, part_numbers, processed_tables):
    """
    Create a comprehensive datasheet with all extracted information
    """
    print("📝 Creating comprehensive datasheet...")
    
    # Create main datasheet structure
    datasheet = {
        'Document_Info': {
            'Title': 'RK73H Series Resistor Data',
            'Extraction_Date': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'Total_Pages': len(extracted_data['text_content']),
            'Total_Tables': len(processed_tables)
        },
        'General_Specifications': specifications,
        'Part_Numbers': part_numbers,
        'Detailed_Tables': processed_tables
    }
    
    return datasheet

def save_to_excel(datasheet, filename="RK73H_Complete_Datasheet.xlsx"):
    """
    Save all extracted data to Excel with multiple sheets
    """
    print(f"💾 Saving to Excel: {filename}")
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        
        # Sheet 1: Document Summary
        doc_info = pd.DataFrame.from_dict(datasheet['Document_Info'], orient='index', columns=['Value'])
        doc_info.to_excel(writer, sheet_name='Document_Info')
        
        # Sheet 2: General Specifications
        if datasheet['General_Specifications']:
            spec_data = []
            for spec_type, values in datasheet['General_Specifications'].items():
                for value in values:
                    spec_data.append({'Specification_Type': spec_type, 'Value': value})
            
            if spec_data:
                spec_df = pd.DataFrame(spec_data)
                spec_df.to_excel(writer, sheet_name='General_Specifications', index=False)
        
        # Sheet 3: Part Numbers
        if datasheet['Part_Numbers']:
            part_df = pd.DataFrame(datasheet['Part_Numbers'], columns=['Part_Number'])
            part_df.to_excel(writer, sheet_name='Part_Numbers', index=False)
        
        # Sheet 4+: Individual Tables
        for i, table_info in enumerate(datasheet['Detailed_Tables']):
            sheet_name = f"Table_Page_{table_info['page']}_{i+1}"
            if len(sheet_name) > 31:  # Excel sheet name limit
                sheet_name = f"Table_{i+1}"
            
            try:
                table_info['dataframe'].to_excel(writer, sheet_name=sheet_name, index=False)
            except Exception as e:
                print(f"   ⚠️ Error saving table {i+1}: {e}")
        
        # Sheet: All Text Content
        text_data = []
        for item in datasheet['_raw_text']:
            text_data.append({
                'Page': item['page'],
                'Content': item['text'][:32000]  # Excel cell limit
            })
        
        if text_data:
            text_df = pd.DataFrame(text_data)
            text_df.to_excel(writer, sheet_name='Raw_Text_Content', index=False)

def main():
    """
    Main function to extract all data from RK73H.pdf
    """
    pdf_path = 'RK73H.pdf'
    
    print("🚀 Starting PDF Data Extraction")
    print("="*50)
    
    try:
        # Step 1: Extract raw data
        extracted_data = extract_pdf_data(pdf_path)
        
        # Step 2: Parse specifications
        specifications = parse_specifications(extracted_data)
        
        # Step 3: Extract part numbers
        part_numbers = extract_part_numbers(extracted_data)
        
        # Step 4: Process tables
        processed_tables = process_tables(extracted_data)
        
        # Step 5: Create comprehensive datasheet
        datasheet = create_comprehensive_datasheet(
            extracted_data, specifications, part_numbers, processed_tables
        )
        
        # Add raw text for reference
        datasheet['_raw_text'] = extracted_data['text_content']
        
        # Step 6: Save to Excel
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"RK73H_Complete_Datasheet_{timestamp}.xlsx"
        save_to_excel(datasheet, output_filename)
        
        # Print summary
        print("\n" + "="*50)
        print("✅ EXTRACTION COMPLETE!")
        print("="*50)
        print(f"📁 Output file: {output_filename}")
        print(f"📄 Pages processed: {len(extracted_data['text_content'])}")
        print(f"📊 Tables found: {len(processed_tables)}")
        print(f"🔢 Part numbers found: {len(part_numbers)}")
        print(f"📋 Specification types: {len(specifications)}")
        
        if part_numbers:
            print(f"\n🔍 Sample part numbers found:")
            for i, part in enumerate(part_numbers[:5]):
                print(f"  {i+1}. {part}")
            if len(part_numbers) > 5:
                print(f"  ... and {len(part_numbers) - 5} more")
        
        if specifications:
            print(f"\n📋 Specifications extracted:")
            for spec_type, values in specifications.items():
                print(f"  • {spec_type}: {len(values)} entries")
        
        return output_filename
        
    except FileNotFoundError:
        print(f"❌ Error: PDF file '{pdf_path}' not found!")
        print("   Make sure the file exists in the current directory.")
        return None
    
    except Exception as e:
        print(f"❌ Error during extraction: {e}")
        return None

if __name__ == "__main__":
    main()
