import re
from pypdf import PdfReader 
import csv
import pandas as pd
import os
import shutil


def start(model, year, file_path):
    """
    Starts the process by reading a file and determining whether to use the spec_method or quote_method.

    Parameters:
    model (str): The model of the item.
    year (int): The year associated with the file.
    file_path (str): The path to the PDF file.

    Returns:
    None
    """
    page_text = read(file_path)
    stripped_text = page_text.lower().replace(" ", "")
    if("quotation" in stripped_text):
        quote_method(file_path, model, year)
    elif ("specificationproposal" in stripped_text):
        spec_method(file_path, model, year)
    else:
        print("UNable to Identify File Types")




def read(file_path):
    """
    Reads the text from a PDF file and returns the extracted content.

    Parameters:
    file_path (str): The path to the PDF file.

    Returns:
    str: Extracted text from the PDF, or an empty string if an error occurs.
    """
    try:
        reader = PdfReader(file_path)
        num_pages = len(reader.pages)
        page_text = ""
        for i in range(num_pages):
            page = reader.pages[i]
            text = page.extract_text()
            if text:
                page_text += text
            else:
                print(f"Warning: Unable to extract text from page {i + 1} in {os.path.basename(file_path)}")
        return page_text
    except FileNotFoundError:
        print(f"Error: The file {os.path.basename(file_path)} was not found.")
    except Exception as e:
        print(f"Error processing file {os.path.basename(file_path)}: {e}")
    return ""


def crop_pdf(input_string, keyword1, keyword2):
    """
    Crops text between two specified keywords from a string.

    Parameters:
    input_string (str): The input text to be cropped.
    keyword1 (str): The starting keyword for cropping.
    keyword2 (str): The ending keyword for cropping.

    Returns:
    str: The cropped text, or None if keywords are not found or in the wrong order.
    """
    start_index = input_string.find(keyword1)
    end_index = input_string.find(keyword2)
    if start_index == -1 or end_index == -1 or start_index >= end_index:
        return None
    start_index += len(keyword1)
    result = input_string[start_index:end_index].strip()
    return result


def crop_pdf2(input_string, keyword, n=None):
    """
    Crops text around a specific keyword and returns a certain number of lines after the keyword.
    If n is not provided, it returns all remaining lines after the keyword.

    Parameters:
    input_string (str): The input text to search within.
    keyword (str): The keyword to locate in the text.
    n (int, optional): The number of lines to include after the keyword. Defaults to None, which returns all remaining lines.

    Returns:
    str: Cropped text containing the keyword and subsequent lines, or None if the keyword is not found.
    """
    lines = input_string.split('\n')
    result_lines = []
    for i, line in enumerate(lines):
        if keyword in line:
            if n is None:
                result_lines = lines[i:]  
            else:
                result_lines = lines[i:i + n + 1] 
            break
    if not result_lines:
        return None
    result = '\n'.join(result_lines)
    return result


def cleaning2(text, keyword1, keyword2, keyword3):
    """
    Cleans text by removing unwanted lines between specific keywords.

    Parameters:
    text (str): The text to be cleaned.
    keyword1 (str): The first keyword to search for.
    keyword2 (str): The keyword to be removed.
    keyword3 (str): The fallback keyword to search for if keyword1 is not present.

    Returns:
    str: The cleaned text with unnecessary lines and keywords removed.
    """
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


def remove_empty(text):
    """
    Removes empty lines from a block of text.

    Parameters:
    text (str): The input text.

    Returns:
    str: The text with all empty lines removed.
    """
    lines = text.split('\n')
    non_empty_lines = [line for line in lines if line.strip() != '']
    new_text = '\n'.join(non_empty_lines)
    return new_text


def join_gaps(text, predefined_array):
    """
    Joins lines that should belong together based on a pattern or a predefined array.

    Parameters:
    text (str): The input text to process.
    predefined_array (list of str): Array of predefined patterns to keep separate.

    Returns:
    str: The processed text with lines joined based on the pattern or predefined array.
    """
    pattern = re.compile(r'^\s*[A-Za-z0-9]{3}-[A-Za-z0-9]{3}')
    lines = text.split('\n')
    processed_lines = []
    for line in lines:
        if pattern.match(line) or line.strip() in predefined_array:
            processed_lines.append(line)
        else:
            if processed_lines:
                processed_lines[-1] += '' + line.strip()
            else:
                processed_lines.append(line.strip())
    return '\n'.join(processed_lines)


def add_heading(text, predefined_array):
    """
    Adds headings to text lines based on predefined keywords.

    Parameters:
    text (str): The input text.
    predefined_array (list of str): Array of predefined headings.

    Returns:
    str: The text with headings added to the corresponding lines.
    """
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


def remove_colon_space(text):
    """
    Removes spaces after colons in a block of text.

    Parameters:
    text (str): The input text to process.

    Returns:
    str: The text with spaces after colons removed.
    """
    processed_lines = []
    for line in text.splitlines():
        processed_line = []
        i = 0
        while i < len(line):
            char = line[i]
            if char == ':' and i + 1 < len(line) and line[i + 1] == ' ':
                processed_line.append(char)
                i += 2  # Increment by 2 to skip the space after the colon
            else:
                processed_line.append(char)
                i += 1
        processed_lines.append(''.join(processed_line))
    return '\n'.join(processed_lines)


def extract(text, headings, model, year, file_path):
    """
    Extracts structured data from text based on a pattern and headings.

    Parameters:
    text (str): The input text.
    headings (list of str): A list of valid headings to look for.
    model (str): The model of the item.
    year (int): The year associated with the data.
    file_path (str): The file path of the document being processed.

    Returns:
    list: A list of extracted data, where each item is a list containing the extracted fields.
    """
    pattern = re.compile(r'^\s*(.*?)\s+([A-Za-z0-9]{3}\s*-\s*[A-Za-z0-9]{3})')
    extracted_data = []
    for line in text.splitlines():
        match = pattern.match(line)
        if not match:
            continue
        keyword = match.group(1)
        xxx_xxx = match.group(2)
        if keyword not in headings:
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
        extracted_data.append([keyword, xxx_xxx, large_string, number1, number2, extra, year, model, file_path])
    return extracted_data


def quote_extract(text, model, year, file_path):
    text = crop_pdf2(text, "VEHICLE PRICE")
    match = re.search(r'\((\d+)\)', text)
    units = int(match.group(1)) if match else None
    
    text = re.sub(r'TOTAL # OF UNITS\s*\(.*?\)\s*', '', text)
    text = re.sub(r'\(.*?\)', '', text)
    
    lines = text.splitlines()
    extracted_data = []
    
    pattern = re.compile(r'(.+?)\s*\$\s*([\d,]+)\s*\$\s*([\d,]+)')
    
    for line in lines:
        match = pattern.search(line)
        if match:
            heading = match.group(1).strip()  
            num1 = match.group(2).replace(',', '') 
            num2 = match.group(3).replace(',', '')  
        
            extracted_data.append([heading, units, num1, num2, year, model, file_path])
    
    return extracted_data

def spec_method(file_path, model, year):
    """
    Processes a specification PDF by extracting and cleaning data, and generating an Excel sheet.

    Parameters:
    file_path (str): The path to the PDF file.
    model (str): The model name.
    year (int): The year of the data.

    Returns:
    None
    """
    original_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputs')
    page_text = read(file_path)  
    keyword1 = "SP E C I F I C A T I O N  PR O P O S A L"
    keyword2 = "T O T A L  V E H I C L E  S U M M A R Y"
    
    headings = [ 
        "Price Level", "Data Version", "Interior Convenience/Driver Retention Package", "Vehicle Configuration",
        "General Service", "Truck Service", "Engine", "Electronic Parameters", "Engine Equipment", "Transmission",
        "Transmission Equipment", "Front Axle and Equipment", "Front Suspension", "Rear Axle and Equipment",
        "Rear Suspension", "Brake System", "Trailer Connections", "Wheelbase & Frame", "Chassis Equipment",
        "Fuel Tanks", "Tires", "Hubs", "Wheels", "Cab Exterior", "Cab Interior", "Instruments & Controls", "Design",
        "Color", "Certification / Compliance", "Secondary Factory Options", "Sales Programs"
    ]

    page_text_cropped = crop_pdf(page_text, keyword1, keyword2)  
    if page_text_cropped is None: 
        print(file_path)
        return 0

    page_text_cleaned = cleaning2(page_text_cropped, "Retail Price", "Prepared for:", "Rear")
    page_text_cleaned = remove_empty(page_text_cleaned)
    page_text_cleaned = join_gaps(page_text_cleaned, headings)
    page_text_cleaned = add_heading(page_text_cleaned, headings)
    page_text_cleaned = remove_colon_space(page_text_cleaned)

    final_array = extract(page_text_cleaned, headings, model, year, file_path)
    warranty = warranty_extraction(file_path, model, year)
    final_array += warranty

    final_csv = pd.DataFrame(final_array)
    final_csv.columns = [
        "Heading", "Data Code", "Description", "Weight Front", "Weight Rear", "Retail Price", "Year", "Work Order", 
        "File Path"
    ]

    with pd.ExcelWriter('Spec.xlsx') as writer:
        final_csv.to_excel(writer, index=False)

    destination_dir = os.path.join(original_path, year, model)
    os.makedirs(destination_dir, exist_ok=True)
    shutil.move('Spec.xlsx', os.path.join(destination_dir, 'Spec.xlsx'))

    weights_summary(file_path, model, year)


def quote_method(file_path, model, year):
    """
    Processes a quote PDF by extracting relevant data and generating an Excel sheet.

    Parameters:
    file_path (str): The path to the PDF file.
    model (str): The model name.
    year (int): The year of the data.

    Returns:
    None
    """
    original_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputs')
    page_text = read(file_path)
    
    final_array = quote_extract(page_text, model, year, file_path)
    final_csv = pd.DataFrame(final_array)
    
    if final_csv.empty:
        return 0
    
    #final_csv = final_csv.drop(final_csv.columns[0], axis=1)
    final_csv.columns = ["Line Items", "Number of Units", "Price per Unit", "Total Price", "Year", "Work Order", "File Path"]
    with pd.ExcelWriter('Quote.xlsx') as writer:
        final_csv.to_excel(writer, index=False)
    destination_dir = os.path.join(original_path, year, model)
    os.makedirs(destination_dir, exist_ok=True)
    shutil.move('Quote.xlsx', os.path.join(destination_dir, 'Quote.xlsx'))


def warranty_extraction(file_path, model, year):
    """
    Extracts warranty data from a PDF file.

    Parameters:
    file_path (str): The path to the PDF file.
    model (str): The model name.
    year (int): The year of the data.

    Returns:
    list: Extracted warranty data.
    """
    page_text = read(file_path)
    cropped_text = crop_pdf2(page_text, "Extended Warranty", 5)
    
    if cropped_text is None:
        print(page_text)

    lines = cropped_text.split('\n')
    processed_lines = ["Extended Warranty " + line.strip() for line in lines]
    final_text = '\n'.join(processed_lines)

    headings = ["Extended Warranty"]
    final_array = extract(final_text, headings, model, year, file_path)

    return final_array


def weights_summary(file_path, model, year):
    """
    Extracts weight summary data from a PDF file and generates an Excel sheet.

    Parameters:
    file_path (str): The path to the PDF file.
    model (str): The model name.
    year (int): The year of the data.

    Returns:
    None
    """
    original_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputs')
    page_text = read(file_path)

    headings = ["Factory Weight", "Dealer Installed Options", "Total Weight"]
    cropped_text = crop_pdf(page_text, "T O T A L  V E H I C L E  S U M M A R Y", "I T E M S  N O T  I N C L U D E D  I N  A D J U S T E D  L I S T")

    if cropped_text is None:
        cropped_text = crop_pdf(page_text, "T O T A L  V E H I C L E  S U M M A R Y", "Extended Warranty")
        if cropped_text is None:
            print(file_path)
            return 0

    cleaned_text = cleaning2(cropped_text, "Rear Total", "Adjusted List Price", "")
    pattern = re.compile(r'\b(\d+)\s+lbs\b')
    extracted_data = []

    for line in cleaned_text.splitlines():
        line = line.strip()
        heading = None

        for h in headings:
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

    final_csv = pd.DataFrame(extracted_data)
    final_csv.columns = ["Headings","Weight Front", "Weight Rear", "Total Weight","Year", "Model"]

    with pd.ExcelWriter('WeightSummary.xlsx') as writer:
        final_csv.to_excel(writer, index=False)

    destination_dir = os.path.join(original_path, year, model)
    os.makedirs(destination_dir, exist_ok=True)
    shutil.move('WeightSummary.xlsx', os.path.join(destination_dir, 'WeightSummary.xlsx'))

def get_input_directory(input_dir=None):
    """
    Returns the input directory path. If input_dir is not specified, it defaults to 'inputs' 
    folder in the script's current directory.

    Parameters:
    input_dir (str): Optional path to the input directory. Defaults to None.

    Returns:
    str: The resolved input directory path.
    """
    if input_dir is None:
        script_dir = os.path.dirname(os.path.abspath(__file__)) 
        input_dir = os.path.join(script_dir, 'inputs')

    return input_dir

    
input_dir = get_input_directory() 
if not os.path.exists(input_dir):
    print(f"Error: The input directory '{input_dir}' does not exist.")
else:
    for name in os.listdir(input_dir):
        path = os.path.join(input_dir, name)
        
        if os.path.isdir(path):  
            year = name
            
            for sub_name in os.listdir(path):
                sub_path = os.path.join(path, sub_name)
                
                if os.path.isdir(sub_path):  
                    model = sub_name 
                    
                    for file in os.listdir(sub_path):
                        file_path = os.path.join(sub_path, file)
                        
                        if os.path.isfile(file_path) and file_path.lower().endswith('.pdf'):
                            start(model, year, file_path) 