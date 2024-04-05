import re
import os
import pandas as pd
import datetime
import pyperclip
import keyboard

current_datetime = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")

def handle_repeatn_input(input_str_or_list):
    if isinstance(input_str_or_list, list):
        repeatn_values = input_str_or_list
    else:
        repeatn_values = input_str_or_list.split(',')

    # Process repeatn values and return them
    if not repeatn_values:
        return ['#NA']  # Handle the case of empty input

    try:
        count_values = [int(repeatn_values[i]) for i in range(0, len(repeatn_values), 2)]  # Extract counts
        values = repeatn_values[1::2]  # Slice to get every second element starting from index 1

        if len(values) % 2 != 0:
            return ['#NA']  # Invalid repeatn values

        repeated_values = []
        for i in range(len(values)):
            value_to_repeat = values[i]
            repeat_count = count_values[i]
            repeated_values.extend([value_to_repeat] * repeat_count)

        if len(repeated_values) < sum(count_values):
            print(f"Warning: Not enough repeatn values provided.")
        
        return repeated_values  # Return the correct repeated values

    except ValueError:
        return ['#NA']  # Invalid repeatn format

# Function to crop a variable to a specified number of digits
def crop_variable(variable, crop_position, num_digits):
    if not isinstance(variable, (str, list)):
        return "#NA"
    
    if crop_position == 'start':
        return variable[num_digits:] if isinstance(variable, str) else variable
    elif crop_position == 'end':
        return variable[:-num_digits] if isinstance(variable, str) else variable
    else:
        return variable

    
def replace_variables(template, variable_values, left_delimiter, right_delimiter):
    for variable, value in variable_values.items():
        placeholder = f"{left_delimiter}{variable}{right_delimiter}"
        template = template.replace(placeholder, str(value))
    return template

# Function to save generated strings to a text file
def save_to_txt_file(generated_strings, txt_file_path, separator):
    try:
        with open(txt_file_path, 'w') as file:
            file.write(separator.join(generated_strings))
        print(f"Results saved to '{txt_file_path}' (txt).")
    except Exception as e:
        print(f"Error saving to text file: {e}")

# Function to create a new Excel file and save generated strings
def create_new_excel_file(generated_strings, excel_file_path, sheet_name):
    try:
        # Create a DataFrame from the generated strings
        df = pd.DataFrame(generated_strings, columns=['Generated Strings'])
        
        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"Results saved to a new Excel file '{excel_file_path}' in a new sheet '{sheet_name}'.")
    except Exception as e:
        print(f"Error saving to Excel file: {e}")

# Function to save generated strings to a text file
def save_to_text_file(generated_strings, output_file_path, separator):
    try:
        with open(output_file_path, 'w') as file:
            if separator == 'enter':
                separator = '\n'  # If the separator is 'enter', use a newline character.
            file.write(separator.join(generated_strings))
        print(f"Results saved to '{output_file_path}'.")
    except Exception as e:
        print(f"Error saving to text file: {e}")

# Function to save generated strings to an Excel file
def save_to_excel_file(generated_strings, excel_file_path, sheet_name):
    try:
        # Create a DataFrame from the generated strings
        df = pd.DataFrame(generated_strings, columns=['Generated Strings'])
        
        # Choose an engine that supports creating a new sheet
        engine = 'openpyxl'
        
        with pd.ExcelWriter(excel_file_path, engine=engine, mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"Results saved to '{excel_file_path}' (Excel) in a new sheet '{sheet_name}'.")
    except Exception as e:
        print(f"Error saving to Excel file: {e}")

def generate_strings(template, variable_values, num_strings, left_delimiter, right_delimiter):
    generated_strings = []
    for i in range(num_strings):
        generated_variable_values = {}
        for variable, values in variable_values.items():
            if isinstance(values, list):
                value_index = i % len(values)
                generated_variable_values[variable] = values[value_index]
            elif variable.endswith("_lookup"):
                # Handle lookup operations
                lookup_variable_name = variable
                base_variable_name = variable.replace("_lookup", "")
                lookup_values = variable_values.get(f"{base_variable_name}_lookup_values", [])
                index_array = variable_values.get(f"{base_variable_name}_index", [])
                output_array = variable_values.get(f"{base_variable_name}_output", [])

                try:
                    lookup_result = perform_lookup(lookup_values, index_array, output_array)
                    generated_variable_values[base_variable_name] = lookup_result
                except Exception as e:
                    print(f"Error during lookup for '{base_variable_name}': {str(e)}")

                # Ensure that there is a count followed by values to repeat
                print("Values:"+ repeatn_values)
                repeatn_count = repeatn_values[0]
                repeatn_values = repeatn_values[1:]
                print("Count:"+ repeatn_count)
                print("Data:"+ repeatn_values)

                if not repeatn_count.isdigit():
                    print(f"Error: Invalid repeatn count for '{variable}'.")
                    generated_strings.append("#NA")
                    continue

                repeatn_count = int(repeatn_count)

                if len(repeatn_values) % 2 != 0:
                    print(f"Error: Invalid repeatn values for '{variable}'.")
                    generated_strings.append("#NA")
                    continue

                repeated_values = []
                for j in range(len(repeatn_values) // 2):
                    value_to_repeat = repeatn_values[j * 2]
                    repeat_count = int(repeatn_values[j * 2 + 1])
                    repeated_values.extend([value_to_repeat] * repeat_count)

                if len(repeated_values) < repeatn_count:
                    print(f"Warning: Not enough repeatn values for '{variable}'.")
                generated_variable_values[variable] = repeated_values[:repeatn_count]
            else:
                generated_variable_values[variable] = values

        generated_string = replace_variables(template, generated_variable_values, left_delimiter, right_delimiter)
        print(generated_string)
        generated_strings.append(generated_string)

    return generated_strings


# Function to read variable values from an Excel file
def read_values_from_excel(excel_file_path, sheet_name, column_name_or_ref, has_headers):
    try:
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=0 if has_headers else None)
        if has_headers:
            if column_name_or_ref in df.columns:
                values = df[column_name_or_ref].tolist()
                return values
            else:
                print(f"Error: Column '{column_name_or_ref}' not found in the Excel sheet '{sheet_name}'.")
        else:
            if isinstance(column_name_or_ref, str) and column_name_or_ref.isalpha() and len(column_name_or_ref) == 1:
                column_index = ord(column_name_or_ref.upper()) - ord('A')
                if 0 <= column_index < len(df.columns):
                    values = df.iloc[:, column_index].tolist()
                    return values
                else:
                    print(f"Error: Invalid column reference '{column_name_or_ref}' in Excel.")
            else:
                print(f"Error: Invalid column reference '{column_name_or_ref}' in Excel.")
        return ['#NA']
    except FileNotFoundError:
        print(f"Error: Excel file not found at '{excel_file_path}'.")
        return ['#NA']
    except Exception as e:
        print(f"Error reading values from Excel: {e}")
        return ['#NA']

def read_values_from_txt_file(txt_file_path, separator):
    try:
        with open(txt_file_path, 'r') as file:
            # Read the file and split using the specified separator
            values = file.read().split(separator)
        return values
    except FileNotFoundError:
        print(f"Error: Text file not found at '{txt_file_path}'.")
        return ['#NA']
    except Exception as e:
        print(f"Error reading values from text file: {e}")
        return ['#NA']


# Function to read sequential values with a specified number of decimals
def read_values_sequential(start, step, order, count, decimals):
    try:
        values = []
        current_value = start
        for _ in range(count):
            formatted_value = f"{current_value:.{decimals}f}"
            values.append(formatted_value)
            
            if order == 'crescent':
                current_value += step
            else:
                current_value -= step
        return values
    except Exception as e:
        print(f"Error generating sequential values: {str(e)}")
        return ['#NA']

def perform_lookup(lookup_values, index_array, output_array):
    lookup_result = []

    for lookup_value in lookup_values:
        found = False  # Flag to check if a match is found

        for i, index_value in enumerate(index_array):
            if isinstance(lookup_value, (int, float)) or isinstance(index_value, (int, float)):
                # If either lookup_value or index_value is a number, compare numerically
                if abs(float(lookup_value) - float(index_value)) < 1e-9:
                    lookup_result.append(output_array[i])
                    found = True
                    break
            elif lookup_value == index_value:
                lookup_result.append(output_array[i])
                found = True
                break

        if not found:
            lookup_result.append('#NA')

    return lookup_result

def rename_files_in_folder(folder_path, generated_strings):
    if not os.path.isdir(folder_path):
        print("Error: The provided path is not a valid directory.")
        return
    
    # Get the list of files in the folder
    files_in_folder = os.listdir(folder_path)
    files_in_folder.sort()  # Sort the files alphabetically
    
    num_files = len(files_in_folder)
    num_strings = len(generated_strings)
    
    print(f"Number of files in the folder: {num_files}")
    print(f"Number of generated strings: {num_strings}")
    
    if num_files == 0:
        print("Warning: There are no files in the folder to rename.")
    elif num_files < num_strings:
        print("Warning: There are more generated strings than files in the folder. Not all strings will be used for renaming.")
        
        for i in range(num_files):
            if i < num_strings:
                new_name = os.path.join(folder_path, f"{generated_strings[i]}{os.path.splitext(files_in_folder[i])[1]}")
                os.rename(os.path.join(folder_path, files_in_folder[i]), new_name)
                print(f"Renamed '{os.path.basename(files_in_folder[i])}' to '{os.path.basename(new_name)}'")
            else:
                print(f"No more generated strings to use.")
                break
    else:
        for i in range(num_files):
            new_name = os.path.join(folder_path, f"{generated_strings[i]}{os.path.splitext(files_in_folder[i])[1]}")
            os.rename(os.path.join(folder_path, files_in_folder[i]), new_name)
            print(f"Renamed '{os.path.basename(files_in_folder[i])}' to '{os.path.basename(new_name)}'")
        
        if num_strings > num_files:
            print("Warning: Not all generated strings were used for renaming.")


# Get the current date and time for file naming
current_datetime = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")

# Prompt the user for advanced features
use_advanced_features = input("Do you want to add advanced features? (y/n): ").strip().lower() == 'y'

# Initialize variables for default values
default_decimals = 0
default_step = 1
default_crescent = True
default_total_integer_digits = 1
left_delimiter = '['
right_delimiter = ']'

# Prompt the user for left and right delimiters
if use_advanced_features:
    left_delimiter = input("Enter the left delimiter for variables (default is '['): ").strip() or '['
    right_delimiter = input("Enter the right delimiter for variables (default is ']'): ").strip() or ']'

print_enabled = True
# Prompt the user to disable printing if advanced features are on
if use_advanced_features:
    disable_print = input("Disable printing of array values? (y/n): ").strip().lower()
    if disable_print == 'y':
        print_enabled = False

# Prompt the user for the template
template = input("Enter the template string: ")


# Extract variable names from the template
variable_pattern = re.escape(left_delimiter) + r'([^' + re.escape(right_delimiter) + r']+)' + re.escape(right_delimiter)
variable_names = re.findall(variable_pattern, template)

# Initialize a dictionary to store variable values (arrays, single values, txt file paths, or Excel)
variable_values = {}
for variable in variable_names:
    input_type = input(f"Enter input type for '{variable}' (txt/excel/manual/lookup/sequential/repeatn): ").strip().lower()
    
    # Initialize values with a default value
    values = None
    
    if input_type == 'txt':
        if use_advanced_features:
            separator = input(f"Enter the separator for '{variable}' (',' for comma, ' ' for space, 'enter' for paragraphs): ").strip()
            separator = '\n' if separator.lower() == 'enter' else separator
        else:
            separator = ','
        txt_file_path = input(f"Enter path to txt file for '{variable}': ").strip()
        txt_file_path = os.path.normpath(txt_file_path.strip('"'))
        values = read_values_from_txt_file(txt_file_path, separator)
        
    elif input_type == 'excel':
        excel_file_path = input(f"Enter path to Excel file for '{variable}': ").strip()
        excel_file_path = os.path.normpath(excel_file_path.strip('"'))  # Normalize path for compatibility
        sheet_name = input(f"Enter Excel sheet name for '{variable}': ").strip()
        has_headers = input("Does the sheet have headers? (y/n): ").strip().lower() == 'y'
        
        if has_headers:
            column_name_or_ref = input(f"Enter Excel column name for '{variable}': ").strip()
        else:
            column_name_or_ref = input(f"Enter Excel column reference for '{variable}' (e.g., 'A' or 'B'): ").strip()
        
        values = read_values_from_excel(excel_file_path, sheet_name, column_name_or_ref, has_headers)
    elif input_type == 'sequential':
        print(f"Setting up '{variable}' as a sequential operation.")
        
        if use_advanced_features:
            try:
                decimals = int(input(f"Enter the number of decimals for '{variable}' (advanced): "))
                step = float(input(f"Enter the step size for '{variable}' (advanced): "))
                order = input(f"Crescent order for '{variable}'? (y/n, advanced): ").strip().lower()
                crescent = order == 'y'
                total_integer_digits = int(input(f"Enter the total integer digits for '{variable}': "))
            except ValueError:
                print("Invalid input. Please enter valid numeric values.")
                decimals = default_decimals
                step = default_step
                crescent = default_crescent
                total_integer_digits = default_total_integer_digits
        else:
            # Use default values
            decimals = default_decimals
            step = default_step
            crescent = default_crescent
        
        try:
            start = float(input(f"Enter the starting number for '{variable}': "))
            count = int(input(f"Enter the number of values to generate for '{variable}' (enter a large number for no end): "))
            
            values = read_values_sequential(start, step, 'crescent' if crescent else 'decreasing', count, decimals)
            if use_advanced_features:
                values = [f"{int(value.split('.')[0]):0{total_integer_digits}d}.{value.split('.')[1]}" if '.' in value else f"{int(value):0{total_integer_digits}d}" for value in values]

        except ValueError:
            print("Invalid input. Please enter valid numeric values.")
            values = ['#NA']
    elif input_type == 'lookup':
        print(f"Setting up '{variable}' as a lookup operation.")
        
        # Retrieve the lookup values
        lookup_input_type = input(f"Enter input type for lookup values' (txt/excel/manual/sequential): ").strip().lower()
        if lookup_input_type == 'txt':
            if use_advanced_features:
                separator = input(f"Enter the separator for lookup values of '{variable}' (',' for comma, ' ' for space, 'enter' for paragraphs): ").strip()
                separator = '\n' if separator.lower() == 'enter' else separator
            else:
                separator = ','

            txt_file_path = input(f"Enter path to txt file for lookup values of '{variable}': ").strip()
            txt_file_path = os.path.normpath(txt_file_path.strip('"'))
            lookup_values = read_values_from_txt_file(txt_file_path, separator)
        elif lookup_input_type == 'excel':
            excel_file_path = input(f"Enter path to Excel file for lookup values': ").strip()
            excel_file_path = os.path.normpath(excel_file_path.strip('"'))
            sheet_name = input(f"Enter Excel sheet name for lookup values': ").strip()
            has_headers = input("Does the sheet have headers? (y/n): ").strip().lower() == 'y'
            
            if has_headers:
                column_name_or_ref = input(f"Enter Excel column name for lookup values': ").strip()
            else:
                column_name_or_ref = input(f"Enter Excel column reference for lookup values' (e.g., 'A' or 'B'): ").strip()
            
            lookup_values = read_values_from_excel(excel_file_path, sheet_name, column_name_or_ref, has_headers)
        elif lookup_input_type == 'sequential':
            print(f"Setting up '{variable}' as a sequential operation for lookup values.")
            
            if use_advanced_features:
                try:
                    decimals = int(input(f"Enter the number of decimals for lookup values (advanced): "))
                    step = float(input(f"Enter the step size for lookup values (advanced): "))
                    order = input(f"Crescent order for lookup values? (y/n, advanced): ").strip().lower()
                    crescent = order == 'y'
                    total_integer_digits = int(input(f"Enter the total integer digits for '{variable}': "))
                except ValueError:
                    print("Invalid input. Please enter valid numeric values.")
                    decimals = default_decimals
                    step = default_step
                    crescent = default_crescent
                    total_integer_digits = default_total_integer_digits
            else:
                # Use default values
                decimals = default_decimals
                step = default_step
                crescent = default_crescent
                total_integer_digits = default_total_integer_digits
            
            try:
                start = float(input(f"Enter the starting number for lookup values: "))
                count = int(input(f"Enter the number of values to generate for lookup values (enter a large number for no end): "))
                
                lookup_values = read_values_sequential(start, step, 'crescent' if crescent else 'decreasing', count, decimals)
                if use_advanced_features:
                    lookup_values = [f"{int(value.split('.')[0]):0{total_integer_digits}d}.{value.split('.')[1]}" if '.' in value else f"{int(value):0{total_integer_digits}d}" for value in lookup_values]
            except ValueError:
                print("Invalid input. Please enter valid numeric values.")
                lookup_values = ['#NA']

        else:
            lookup_values_input = input(f"Enter lookup values (comma-separated): ")
            lookup_values = lookup_values_input.split(',')
        if print_enabled:
            print('Lookup Values:\n' + str(lookup_values))

        # Retrieve the index array
        index_input_type = input(f"Enter input type for index array (txt/excel/manual/sequential): ").strip().lower()
        if index_input_type == 'txt':
            if use_advanced_features:
                separator = input(f"Enter the separator for index array of '{variable}' (',' for comma, ' ' for space, 'enter' for paragraphs): ").strip()
                separator = '\n' if separator.lower() == 'enter' else separator
            else:
                separator = ','

            txt_file_path = input(f"Enter path to txt file for index array of '{variable}': ").strip()
            txt_file_path = os.path.normpath(txt_file_path.strip('"'))
            index_array = read_values_from_txt_file(txt_file_path, separator)
        elif index_input_type == 'excel':
            excel_file_path = input("Enter path to Excel file for index array: ").strip()
            excel_file_path = os.path.normpath(excel_file_path.strip('"'))
            sheet_name = input("Enter Excel sheet name for index array: ").strip()
            has_headers = input("Does the sheet have headers? (y/n): ").strip().lower() == 'y'
            
            if has_headers:
                column_name_or_ref = input("Enter Excel column name for index array: ").strip()
            else:
                column_name_or_ref = input("Enter Excel column reference for index array (e.g., 'A' or 'B'): ").strip()
            
            index_array = read_values_from_excel(excel_file_path, sheet_name, column_name_or_ref, has_headers)
        elif index_input_type == 'sequential':
            print("Setting up index array as a sequential operation.")
            
            if use_advanced_features:
                try:
                    decimals = int(input(f"Enter the number of decimals for index array (advanced): "))
                    step = float(input(f"Enter the step size for index array (advanced): "))
                    order = input(f"Crescent order for index array? (y/n, advanced): ").strip().lower()
                    crescent = order == 'y'
                    total_integer_digits = int(input(f"Enter the total integer digits for '{variable}': "))
                except ValueError:
                    print("Invalid input. Please enter valid numeric values.")
                    decimals = default_decimals
                    step = default_step
                    crescent = default_crescent
                    total_integer_digits = default_total_integer_digits
            else:
                # Use default values
                decimals = default_decimals
                step = default_step
                crescent = default_crescent
                total_integer_digits = default_total_integer_digits
            
            try:
                start = float(input("Enter the starting number for index array: "))
                count = int(input("Enter the number of values to generate for index array (enter a large number for no end): "))
                
                index_array = read_values_sequential(start, step, 'crescent' if crescent else 'decreasing', count, decimals)
                if use_advanced_features:
                    index_array = [f"{int(value.split('.')[0]):0{total_integer_digits}d}.{value.split('.')[1]}" if '.' in value else f"{int(value):0{total_integer_digits}d}" for value in index_array]
            except ValueError:
                print("Invalid input. Please enter valid numeric values.")
                index_array = ['#NA']
        else:
            index_input = input("Enter index array values (comma-separated): ")
            index_array = index_input.split(',')
        if print_enabled:
            print('Index Array:\n' + str(index_array))
        
        # Retrieve the output array
        output_input_type = input(f"Enter input type for output array (txt/excel/manual/sequential): ").strip().lower()
        if output_input_type == 'txt':
            if use_advanced_features:
                separator = input(f"Enter the separator for output array of '{variable}' (',' for comma, ' ' for space, 'enter' for paragraphs): ").strip()
                separator = '\n' if separator.lower() == 'enter' else separator
            else:
                separator = ','

            txt_file_path = input(f"Enter path to txt file for output array of '{variable}': ").strip()
            txt_file_path = os.path.normpath(txt_file_path.strip('"'))
            output_array = read_values_from_txt_file(txt_file_path, separator)
        elif output_input_type == 'excel':
            excel_file_path = input("Enter path to Excel file for output array: ").strip()
            excel_file_path = os.path.normpath(excel_file_path.strip('"'))
            sheet_name = input("Enter Excel sheet name for output array: ").strip()
            has_headers = input("Does the sheet have headers? (y/n): ").strip().lower() == 'y'
            
            if has_headers:
                column_name_or_ref = input("Enter Excel column name for output array: ").strip()
            else:
                column_name_or_ref = input("Enter Excel column reference for output array (e.g., 'A' or 'B'): ").strip()
            
            output_array = read_values_from_excel(excel_file_path, sheet_name, column_name_or_ref, has_headers)
        elif output_input_type == 'sequential':
            print("Setting up output array as a sequential operation.")
            
            if use_advanced_features:
                try:
                    decimals = int(input(f"Enter the number of decimals for output array (advanced): "))
                    step = float(input(f"Enter the step size for output array (advanced): "))
                    order = input(f"Crescent order for output array? (y/n, advanced): ").strip().lower()
                    crescent = order == 'y'
                    total_integer_digits = int(input(f"Enter the total integer digits for '{variable}': "))
                except ValueError:
                    print("Invalid input. Please enter valid numeric values.")
                    decimals = default_decimals
                    step = default_step
                    crescent = default_crescent
                    total_integer_digits = default_total_integer_digits
            else:
                # Use default values
                decimals = default_decimals
                step = default_step
                crescent = default_crescent
                total_integer_digits = default_total_integer_digits
            try:
                start = float(input("Enter the starting number for output array: "))
                count = int(input("Enter the number of values to generate for output array (enter a large number for no end): "))
                
                output_array = read_values_sequential(start, step, 'crescent' if crescent else 'decreasing', count, decimals)
                if use_advanced_features:
                    output_array = [f"{int(value.split('.')[0]):0{total_integer_digits}d}.{value.split('.')[1]}" if '.' in value else f"{int(value):0{total_integer_digits}d}" for value in output_array]
            except ValueError:
                print("Invalid input. Please enter valid numeric values.")
                output_array = ['#NA']
        else:
            output_input = input("Enter output array values (comma-separated): ")
            output_array = output_input.split(',')
        
        # Perform the lookup operation
        lookup_result = perform_lookup(lookup_values, index_array, output_array)
        values = lookup_result

    # Within the loop for variable input types
    elif input_type == 'repeatn':
        repeatn_input_type = input(f"Enter input type for repeatn values of '{variable}' (txt/excel/manual): ").strip().lower()

        if repeatn_input_type == 'txt':
            if use_advanced_features:
                separator = input(f"Enter the separator for repeatn values of '{variable}' (',' for comma, ' ' for space, 'enter' for paragraphs): ").strip()
                separator = '\n' if separator.lower() == 'enter' else separator
            else:
                separator = ','

            txt_file_path = input(f"Enter path to txt file for repeatn values of '{variable}': ").strip()
            txt_file_path = os.path.normpath(txt_file_path.strip('"'))
            repeatn_input = read_values_from_txt_file(txt_file_path, separator)
            if print_enabled:
                print('Imported array:\n' + str(values))
        elif repeatn_input_type == 'excel':
            excel_file_path = input(f"Enter path to Excel file for repeatn values of '{variable}': ").strip()
            excel_file_path = os.path.normpath(excel_file_path.strip('"'))
            sheet_name = input(f"Enter Excel sheet name for repeatn values of '{variable}': ").strip()
            has_headers = input(f"Does the sheet have headers? (y/n) for repeatn values of '{variable}': ").strip().lower() == 'y'

            if has_headers:
                column_name_or_ref = input(f"Enter Excel column name for repeatn values of '{variable}': ").strip()
            else:
                column_name_or_ref = input(f"Enter Excel column reference for repeatn values of '{variable}' (e.g., 'A' or 'B'): ").strip()

            repeatn_input = read_values_from_excel(excel_file_path, sheet_name, column_name_or_ref, has_headers)
            if print_enabled:
                print('Imported array:\n' + str(values))
        elif repeatn_input_type == 'manual':
            repeatn_input = input(f"Enter repeatn values for '{variable}' (count followed by values, comma-separated): ").strip()
            repeatn_input = repeatn_input.split(',')
        else:
            repeatn_input = ['#NA']

        values = handle_repeatn_input(repeatn_input)

    else:
        # Manual input as a fallback
        input_value = input(f"Enter value(s) for '{variable}' (comma-separated for multiple values): ")
        values = input_value.split(',')
    
    if use_advanced_features:
        crop_variable_option = input(f"Do you want to crop '{variable}'? (y/n, advanced): ").strip().lower()
        if crop_variable_option == 'y':
            crop_position = input(f"Crop '{variable}' at the start or end? (start/end, advanced): ").strip().lower()
            num_digits = int(input(f"Enter the number of digits to crop for '{variable}' (advanced): "))
            values = [crop_variable(value, crop_position, num_digits) for value in values]

    if print_enabled:
        print(f"'{variable}' values:\n" + str(values))
    variable_values[variable] = values

# Generate strings and save to a file or clipboard
num_strings = int(input("Enter the number of strings to generate: "))

generated_strings = generate_strings(template, variable_values, num_strings, left_delimiter, right_delimiter)

# Copy the generated strings to the clipboard
pyperclip.copy("\n".join(generated_strings))
print("Generated strings copied to clipboard.")

# Prompt the user to save the generated strings to a file or Excel
if use_advanced_features:
    save_option = input("Do you want to save the generated strings to a file? (y/n): ").strip().lower()
    if save_option == 'y':
        output_file_type = input("Enter the output file type (txt/excel): ").strip().lower()
        
        if output_file_type == 'txt':
            # Construct the txt file name based on current date and time
            current_datetime = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
            txt_file_name = f"Strings_{current_datetime}.txt"
            txt_file_path = input("Enter the path to save the txt file: ").strip()
            
            # Remove double quotation marks if they exist around the path
            txt_file_path = txt_file_path.strip('"')
            
            # Check if the provided path is a directory and create the file
            if os.path.isdir(txt_file_path):
                txt_file_path = os.path.join(txt_file_path, txt_file_name)
            else:
                txt_file_path = os.path.normpath(txt_file_path)  # Normalize path for compatibility
            separator = input("Enter the separator to use between words (write 'enter' for a paragraph): ")
            save_to_txt_file(generated_strings, txt_file_path, separator)
        elif output_file_type == 'excel':
            # Construct the excel file name based on current date and time
            current_datetime = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
            excel_file_name = f"Strings_{current_datetime}.xlsx"
            excel_file_path = input("Enter the path to the Excel file: ").strip()
            
            # Remove double quotation marks if they exist around the path
            excel_file_path = excel_file_path.strip('"')
            
            # Check if the provided path is a directory and create the file
            if os.path.isdir(excel_file_path):
                excel_file_path = os.path.join(excel_file_path, excel_file_name)
            else:
                excel_file_path = os.path.normpath(excel_file_path)  # Normalize path for compatibility
            sheet_name = input("Enter the Excel sheet name: ").strip()
            
            # Check if the Excel file already exists and create a new sheet if it doesn't
            if not os.path.exists(excel_file_path):
                create_new_excel_file(generated_strings, excel_file_path, sheet_name)
            else:
                print(f"Warning: A new sheet '{sheet_name}' will be created in the existing Excel file.")
                save_to_excel_file(generated_strings, excel_file_path, sheet_name)
        else:
            print("Invalid output file type. Results not saved.")
    else:
        print("Results not saved.")

if use_advanced_features:
    rename_files_option = input("Do you want to use the generated strings to rename files in a folder? (y/n): ").strip().lower()
    if rename_files_option == 'y':
        folder_path = input("Enter the path to the folder containing the files you want to rename: ").strip('"')
        rename_files_in_folder(folder_path, generated_strings)

print("Press 'Escape' to close the program...")
keyboard.wait("esc")