import re
from pypdf import PdfReader 
import csv
import pandas as pd
import os
import shutil


def Start(model, year, file_path):
    pageText = read(file_path)
    if "SP E C I F I C A T I O N  PR O P O S A L" in pageText:
        SpecMethod(file_path, model, year)
    else:
        QuoteMethod(file_path, model, year)

def read(file_path):
    try:
        reader = PdfReader(file_path) 
        numPages = len(reader.pages)
        pageText = ""
        for i in range(numPages):
            page = reader.pages[i]
            text = page.extract_text()
            if text:
                pageText += text
            else:
                print(f"Warning: Unable to extract text from page {i+1} in {os.path.basename(file_path)}")
        return pageText
    except FileNotFoundError:
        print(f"Error: The file {os.path.basename(file_path)} was not found.")
    except Exception as e:
        print(f"Error processing file {os.path.basename(file_path)}: {e}")
    return ""

def CropPDF(input_string, keyword1, keyword2):
    start_index = input_string.find(keyword1)
    end_index = input_string.find(keyword2)
    if start_index == -1 or end_index == -1 or start_index >= end_index:
        return None
    start_index += len(keyword1)    
    result = input_string[start_index:end_index].strip()
    return result

def CropPDF2(input_string, keyword, n):
    lines = input_string.split('\n') 
    result_lines = []
    for i, line in enumerate(lines):
        if keyword in line:
            result_lines = lines[i:i + n + 1]
            break
    if not result_lines:
        return None
    result = '\n'.join(result_lines)
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

def Extract(text, Headings, model, year, file_path):
    #pattern = re.compile(r'^\s*(.*?)\s+([A-Za-z0-9]{3}-[A-Za-z0-9]{3})')
    pattern = re.compile(r'^\s*(.*?)\s+([A-Za-z0-9]{3}\s*-\s*[A-Za-z0-9]{3})')
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
        extra = 0  
        if len(parts) > 1:
            if parts[1] and parts[1][0] == ' ':
                parts[1] = '0 ' + parts[1][1:]
            numbers_part = parts[1].strip()
            dollar_match = re.search(r'\$\d{1,3}(,\d{3})*(\.\d+)?', numbers_part)
            if dollar_match:
                try:
                    extra = float(dollar_match.group()[1:].replace(',', ''))
                except ValueError:
                    extra = ''
            if dollar_match:
                numbers_part = numbers_part.replace(dollar_match.group(), '')
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
        extracted_data.append([keyword, xxx_xxx, large_string, number1, number2, extra, year, model, file_path])
    return extracted_data

def QuoteExtract(text, Headings, model, year, file_path):
    extracted_data = []
    bracket_value = None
    for line in text.splitlines():
        line = line.strip().replace(',', '')
        heading = None
        for h in Headings:
            if line.startswith(h):
                heading = h
                rest_of_line = line[len(h):].strip()
                break
        if heading is None:
            continue
        rest_of_line = rest_of_line.strip().replace(' ', '')
        bracket_match = re.search(r'\((\d+)\)', rest_of_line)
        if bracket_match:
            bracket_value = bracket_match.group(1)
            rest_of_line = rest_of_line.replace(bracket_match.group(0), '').strip()
        elif bracket_value is None:
            continue 
        number_matches = re.findall(r'\$([\d.]+)', rest_of_line)
        if len(number_matches) < 2:
            continue  
        try:
            number1 = int(float(number_matches[0]))
        except ValueError:
            number1 = 0  
        try:
            number2 = int(float(number_matches[1]))
        except ValueError:
            number2 = 0  
        extracted_data.append([heading, bracket_value, number1, number2, year, model, file_path])
    return extracted_data

def SpecMethod(file_path, model, year):
    original_path = 'C:\\NewOutput'
    pageText12 = read(file_path)
    keyword1 = "SP E C I F I C A T I O N  PR O P O S A L"
    keyword2 ="T O T A L  V E H I C L E  S U M M A R Y"
    Headings = ["Price Level" , "Data Version", "Interior Convenience/Driver Retention Package", "Vehicle Configuration", "General Service", "Truck Service", "Engine", "Electronic Parameters", "Engine Equipment", "Transmission", "Transmission Equipment", "Front Axle and Equipment", "Front Suspension", "Rear Axle and Equipment", "Rear Suspension", "Brake System", "Trailer Connections", "Wheelbase & Frame", "Chassis Equipment", "Fuel Tanks", "Tires", "Hubs", "Wheels",
                "Cab Exterior", "Cab Interior", "Instruments & Controls", "Design", "Color", "Certification / Compliance", "Secondary Factory Options", "Sales Programs"]
    pageText1 = CropPDF(pageText12,keyword1,keyword2)
    if pageText1 == None:
        print(file_path)
        return 0
    pageText1 = Cleaning2(pageText1,"Retail  Price","Prepared for:", "Rear")
    pageText1 = removeEmpty(pageText1)
    pageText1 = JoinGaps(pageText1, Headings)
    pageText1 = AddHeading( pageText1, Headings)
    pageText1 = removeColonSpace(pageText1)
    FinalArray = Extract(pageText1, Headings, model, year, file_path)
    Warranty = WarrantyExtraction(file_path, model, year)
    FinalArray = FinalArray + Warranty
    FinalCSV = pd.DataFrame(FinalArray)
    FinalCSV.columns = ["Heading", "Data Code", "Description", "Weight Front", "Weight Rear", "Retail Price", "Year", "Word Order", "File Path"]
    with pd.ExcelWriter('Spec.xlsx') as writer:
        FinalCSV.to_excel(writer)
    destination_dir = os.path.join(os.path.join(original_path, year),model)
    os.makedirs(destination_dir, exist_ok=True)
    shutil.move('Spec.xlsx', os.path.join(destination_dir, 'Spec.xlsx'))
    weightsumm(file_path, model, year)
    
def QuoteMethod(file_path, model, year):
    original_path = "C:\\NewOutput"
    pageText = read(file_path)
    QuoteHeadings = ["VEHICLE PRICE", "EXTENDED WARRANTY", "CUSTOMER PRICE BEFORE TAX", "BALANCE DUE ", "CUSTOMER PRICE BEFORE TAX", "FEDERAL EXCISE TAX (FET)", "BALANCE DUE" ]
    FinalArray = QuoteExtract(pageText, QuoteHeadings, model, year, file_path)
    FinalCSV = pd.DataFrame(FinalArray)
    if(FinalCSV.shape == (0,0)):
        return 0
    FinalCSV = FinalCSV.drop(FinalCSV.columns[0], axis=1)

    FinalCSV.columns = ["Number of Units", "Price per Unit", "Total Price", "Year", "Work Order", "File Path"]
    with pd.ExcelWriter('Quote.xlsx') as writer:
        FinalCSV.to_excel(writer, index=False)
    destination_dir = os.path.join(os.path.join(original_path, year),model)
    os.makedirs(destination_dir, exist_ok=True)
    shutil.move('Quote.xlsx', os.path.join(destination_dir, 'Quote.xlsx'))

def WarrantyExtraction(file_path, model, year):
    pageText1 = read(file_path)
    pageText = CropPDF2(pageText1,"Extended Warranty",5)
    if (pageText == None):
        print(pageText1)
    lines = pageText.split('\n')
    processed_lines = []
    for line in lines:
        stripped_line = line.strip()
        processed_lines.append("Extended Warranty " + stripped_line)
    pageText = '\n'.join(processed_lines)
    Headings = ["Extended Warranty"]
    FinalArray = Extract(pageText, Headings, model, year, file_path)
    return FinalArray

def weightsumm(file_path, model, year):
    original_path = 'C:\\NewOutput'
    pageText = read(file_path)
    Headings = ["Factory Weight", "Dealer Installed Options", "Total Weight"]
    pageText1 = CropPDF(pageText, "T O T A L  V E H I C L E  S U M M A R Y", "I T E M S  N O T  I N C L U D E D  I N  A D J U S T E D  L I S T")
    if pageText1 == None:
        pageText2 = CropPDF(pageText, "T O T A L  V E H I C L E  S U M M A R Y", "Extended Warranty")
        if pageText2 == None:
            print(file_path)
            return 0
    pageText = Cleaning2(pageText, "Rear  Total", "Adjusted List Price", "")
    pattern = re.compile(r'\b(\d+)\s+lbs\b')
    extracted_data = []
    for line in pageText.splitlines():
        line = line.strip()
        heading = None
        for h in Headings:
            if line.startswith(h):
                heading = h
                rest_of_line = line[len(h):].strip()
                break
        if heading is None:
            continue
        matches = pattern.findall(rest_of_line)
        if len(matches) != 3:
            continue 
        try:
            number1 = int(matches[0])
            number2 = int(matches[1])
            number3 = int(matches[2])
        except ValueError:
            continue 

        extracted_data.append([heading, number1, number2, number3, year, model])
        FinalCSV = pd.DataFrame(extracted_data)
        FinalCSV = FinalCSV.drop(FinalCSV.columns[0], axis=1)
        FinalCSV.columns = ["Weight Front", "Weight Rear", "Total Weight", "Year", "Work Order"]
        with pd.ExcelWriter('WeightSummary.xlsx') as writer:
            FinalCSV.to_excel(writer, index=False)
        destination_dir = os.path.join(os.path.join(original_path, year),model)
        os.makedirs(destination_dir, exist_ok=True)
        shutil.move('WeightSummary.xlsx', os.path.join(destination_dir, 'WeightSummary.xlsx'))

input ="C:\\TerexClone\\Terex\\Freightliner\\inputs"
for name in os.listdir(input):
    path = os.path.join(input, name)
    if os.path.isdir(path):  
        year = name      
        for sub_name in os.listdir(path):
            sub_path = os.path.join(path, sub_name)
            if os.path.isdir(sub_path):  
                model = sub_name 
                for file in os.listdir(sub_path):
                    file_path = os.path.join(sub_path, file)    
                    if os.path.isfile(file_path) and file_path.lower().endswith('.pdf'):
                        file_name = file
                        Start(model, year, file_path) 


