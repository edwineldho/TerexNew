import re
from pypdf import PdfReader 
import csv
import pandas as pd
reader = PdfReader("Commander 4047 4x4 M2-106 2020MY Specs.pdf") 

def CropPDF(input_string, keyword1, keyword2):
    start_index = input_string.find(keyword1)
    end_index = input_string.find(keyword2)
    if start_index == -1 or end_index == -1 or start_index >= end_index:
        return None
    start_index += len(keyword1)    
    result = input_string[start_index:end_index].strip()
    return result


def Cleaning2(text, keyword1, keyword2, keyword3):
    lines = text.split('\n')
    i = 0
    keyword_to_use = keyword1 if keyword1 in text else keyword3

    while i < len(lines):
        if lines[i].strip() == keyword_to_use:
            j = i - 1
            while j >= 0 and keyword2 not in lines[j]:
                j -= 1            
            if j >= 0 and keyword2 in lines[j]:
                lines = lines[:j] + [lines[j].replace(keyword2, '').strip()] + lines[i + 1:]
                i = j 
                continue
        i += 1
    
    new_text = '\n'.join(lines)
    return new_text


def removeEmpty(text):
    lines = text.split('\n')
    non_empty_lines = [line for line in lines if line.strip() != '']
    new_text = '\n'.join(non_empty_lines)
    return new_text

def JoinGaps(text, predefined_array):
    pattern = re.compile(r'^\s*[A-Za-z0-9]{3}-[A-Za-z0-9]{3}')    
    lines = text.split('\n')    
    processed_lines = []
    for line in lines:
        if (pattern.match(line) or line.strip() in predefined_array):
            processed_lines.append(line)
        else:
            if processed_lines:
                processed_lines[-1] += '' + line.strip()
            else:
                processed_lines.append(line.strip())
    return '\n'.join(processed_lines)

def AddHeading(text, predefined_array):
    lines = text.split('\n')
    processed_lines = []
    current_predefined = None
    for line in lines:
        stripped_line = line.strip()
        if stripped_line in predefined_array:
            current_predefined = stripped_line
        else:
            if current_predefined:
                processed_lines.append(f"{current_predefined} {line.strip()}")
            else:
                processed_lines.append(line)
    return '\n'.join(processed_lines)

def removeColonSpace(text):
    processed_lines = []
    for line in text.splitlines():
        processed_line = []
        i = 0
        while i < len(line):
            char = line[i]
            if char == ':' and i + 1 < len(line) and line[i + 1] == ' ':
                processed_line.append(char)
                i += 2 
            else:
                processed_line.append(char)
                i += 1
        processed_lines.append(''.join(processed_line))
    return '\n'.join(processed_lines)

def Extract(text):
    pattern = re.compile(r'^\s*(.*?)\s+([A-Za-z0-9]{3}-[A-Za-z0-9]{3})')
    extracted_data = []
    for line in text.splitlines():
        match = pattern.match(line)
        if not match:
            continue
        keyword = match.group(1)
        xxx_xxx = match.group(2)
        if keyword not in Headings:
            continue
        rest_of_line = line[match.end():].strip()
        parts = rest_of_line.split('  ', 1)  
        large_string = parts[0].strip() if parts else ''
        number1 = 0
        number2 = 0
        extra = ''  # Initialize extra here

        if len(parts) > 1:
            if parts[1] and parts[1][0] == ' ':
                parts[1] = '0 ' + parts[1][1:]
            numbers_part = parts[1].strip()

            # Check for the $ sign, N/C, or STD before splitting
            dollar_match = re.search(r'\$\d{1,3}(,\d{3})*(\.\d+)?', numbers_part)
            nc_match = re.search(r'\bN/C\b', numbers_part)
            std_match = re.search(r'\bSTD\b', numbers_part)

            if dollar_match:
                try:
                    extra = float(dollar_match.group()[1:].replace(',', ''))
                except ValueError:
                    extra = ''
            elif nc_match:
                extra = 'N/C'
            elif std_match:
                extra = 'STD'

            # Remove the extra part from numbers_part to avoid interference
            if dollar_match:
                numbers_part = numbers_part.replace(dollar_match.group(), '')
            elif nc_match:
                numbers_part = numbers_part.replace(nc_match.group(), '')
            elif std_match:
                numbers_part = numbers_part.replace(std_match.group(), '')

            numbers = re.split(r'\s{1,2}', numbers_part.strip())

            if len(numbers) > 0 and numbers[0]: 
                try:
                    number1 = int(float(numbers[0]))
                except ValueError:
                    number1 = 0 

            if len(numbers) > 1 and numbers[1]: 
                try:
                    number2 = int(float(numbers[1]))
                except ValueError:
                    number2 = 0 
            else:
                number2 = 0

        extracted_data.append([keyword, xxx_xxx, large_string, number1, number2, extra])
    
    return extracted_data


numPages = len(reader.pages)
pageText = ""
for i in range (numPages):
    page = reader.pages[i]
    pageText = pageText + page.extract_text()

keyword1 = "SP E C I F I C A T I O N  PR O P O S A L"
keyword2 ="T O T A L  V E H I C L E  S U M M A R Y"
Headings = ["Price Level", "Data Version", "Interior Convenience/Driver Retention Package", "Vehicle Configuration", "General Service", "Truck Service", "Engine", "Electronic Parameters", "Engine Equipment", "Transmission", "Transmission Equipment", "Front Axle and Equipment", "Front Suspension", "Rear Axle and Equipment", "Rear Suspension", "Brake System", "Trailer Connections", "Wheelbase & Frame", "Chassis Equipment", "Fuel Tanks", "Tires", "Hubs", "Wheels",
            "Cab Exterior", "Cab Interior", "Instruments & Controls", "Design", "Color", "Certification / Compliance", "Secondary Factory Options", "Sales Programs"]
pageText1 = CropPDF(pageText,keyword1,keyword2)
pageText1 = Cleaning2(pageText1,"Retail  Price","Prepared for:", "Rear")
pageText1 = removeEmpty(pageText1)
pageText1 = JoinGaps(pageText1, Headings)
pageText1 = AddHeading( pageText1, Headings)
pageText1 = removeColonSpace(pageText1)

print(pageText)

"""
FinalArray = Extract(pageText)

outer_length = len(FinalArray)
FinalCSV = pd.DataFrame(FinalArray)
with pd.ExcelWriter('TestV2.3.xlsx') as writer:
    FinalCSV.to_excel(writer, index=False)"""