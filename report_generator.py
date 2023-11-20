import re
import pandas as pd
import os
import glob

folder_path = 'input'
exclude_elements = 'exclude_elements.txt'
exclude_regex = 'exclude_regex.txt'
column_id = 'TextId'
column_source = '简体中文 zh-CN(4.3-dev)'
column_target = '俄语 ru-RU'
column_output_name = 'Issues found'
report_html_file = 'output.html'


def find_excel_files(folder):
    """
    Returns a list of paths to Excel files (both .xlsx and .xls) in the specified folder.
    """
    # Define the patterns for Excel file extensions
    patterns = ['*.xlsx', '*.xls']

    # List to store the paths of Excel files
    excel_files = []

    # Iterate through each pattern and add matching files to the list
    for pattern in patterns:
        excel_files.extend(glob.glob(os.path.join(folder, pattern)))

    return excel_files


def create_regex_from_list(elements):
    # Escape all special characters in each element
    escaped_elements = [re.escape(element) for element in elements]

    # Join the escaped elements with the pipe character for OR logic
    regex_pattern = '|'.join(escaped_elements)

    return regex_pattern


def create_regex_from_regex_list(elements):
    # Join the escaped elements with the pipe character for OR logic
    regex_pattern = '|'.join(elements)

    return regex_pattern


def read_file_to_list(file_path):
    """
    Reads a text file and returns a list of its lines.
    Each line in the file becomes a new element in the list.
    """
    with open(file_path, 'r', encoding='utf-8') as file:
        # Read lines into a list, stripping newline characters
        lines = [line.strip() for line in file]

    return lines


exclusion_all = fr'{(create_regex_from_regex_list(read_file_to_list(exclude_regex)))}|{(create_regex_from_list(read_file_to_list(exclude_elements)))}'
re_all = re.compile(exclusion_all)


def find_cjk_characters(text):
    regex_cjk_characters = r'[\u4E00-\u9FFF\u3000-\u303F\u3400-\u4DBF]'  # to catch Chinese thing in translation
    cjk_regex = re.compile(regex_cjk_characters)
    matches = cjk_regex.findall(text)
    return matches


def find_latin_letters(text):
    cleaned_text = clean_source_text(text)
    # Adjusted regex to match sequences of Latin letters
    latin_regex = re.compile(r'[a-zA-Z]+')
    matches = latin_regex.findall(cleaned_text)
    return matches


def clean_source_text(text):
    result = re_all.sub('', text)
    return result


def find_in_df(df, latin, chinese):
    parameter_find_latin = latin
    parameter_find_chinese = chinese
    df[column_target].replace('', pd.NA, inplace=True)
    df = df.dropna(subset=[column_target])

    filtered_latin = pd.DataFrame()
    filtered_chinese = pd.DataFrame()

    # latin part
    if parameter_find_latin is True:
        filtered_latin = (df[df[column_target].apply(lambda x: len(find_latin_letters(str(x))) > 0)]).copy(deep=True)
        filtered_latin[column_output_name] = None
        if filtered_latin.empty is False:
            filtered_latin[column_output_name] = filtered_latin[column_target].apply(
                lambda x: find_latin_letters(str(x)))
        else:
            pass

    # chinese part
    if parameter_find_chinese is True:
        filtered_chinese = (df[df[column_target].apply(lambda x: len(find_cjk_characters(str(x))) > 0)]).copy(deep=True)
        filtered_chinese[column_output_name] = None
        if filtered_chinese.empty is False:
            filtered_chinese[column_output_name] = filtered_chinese[column_target].apply(
                lambda x: find_cjk_characters(str(x)))
        else:
            pass

    merged_df = pd.concat([filtered_latin, filtered_chinese], ignore_index=True)
    return merged_df


def load_excel_sheets(folder, source, target, col_id, latin, chinese):
    files = find_excel_files(folder)
    all_sheets_df = {}
    print(f"Loaded excels {files}")
    for file_path in files:
        # Load the Excel file
        xls = pd.ExcelFile(file_path)

        # Get a file-specific prefix to prepend to sheet names
        file_prefix = os.path.splitext(os.path.basename(file_path))[0]

        # Iterate through each sheet in the Excel file
        for sheet_name in xls.sheet_names:
            # Read the specified columns from the sheet into a DataFrame
            df = pd.read_excel(xls, sheet_name=sheet_name, usecols=[source, target, col_id])
            result_df = find_in_df(df, latin, chinese)

            # Rename the sheet to include the file name
            unique_sheet_name = f"{file_prefix} __________________ {sheet_name}"

            # Add the DataFrame to the dictionary if not empty or all null
            if not result_df.empty and not result_df.isnull().all().all():
                all_sheets_df[unique_sheet_name] = result_df

    return all_sheets_df


def merge_adjacent_spans(html_string):
    """
    Merges adjacent <span> elements with the same style in an HTML-like string.
    """
    # Regular expression to find spans with styles
    span_pattern = r'<span style=\'(.*?)\'>(.*?)<\/span>'

    # Function to replace consecutive spans with the same style
    def replace(match):
        spans = re.findall(span_pattern, match.group())
        if not spans:
            return match.group()

        current_style = spans[0][0]
        merged_content = ''

        for style, content in spans:
            if style == current_style:
                merged_content += content
            else:
                # Close the previous span and start a new one
                merged_content += f"</span><span style='{style}'>{content}"
                current_style = style

        return f"<span style='{current_style}'>{merged_content}</span>"

    # Regular expression to find consecutive spans
    consecutive_spans_pattern = r'(<span style=\'[^\']*\'>.*?</span>)+'

    # Replace the matches in the string
    return re.sub(consecutive_spans_pattern, replace, html_string, flags=re.DOTALL)


def highlight_letters(string, matches):
    """
    Highlights the specified letters in the string, but checks each match to see if it
    should be excluded based on the regex_exclude pattern.
    """
    # Combine the letters into a regex pattern
    color = 'red'
    letters_pattern = '[' + ''.join(matches) + ']'
    # Define a regex pattern to find the letters and the excluded patterns
    pattern = fr'(?:{letters_pattern})|(?:{exclusion_all})'

    # Function to replace the match
    def replace(match):
        # If the match is part of the excluded pattern, return it unchanged
        if re.fullmatch(exclusion_all, match.group()):
            return match.group()
        # Otherwise, return the highlighted version
        return f"<span style='color: {color};'>{match.group()}</span>"

    # Replace the matches in the string
    result = re.sub(pattern, replace, string)
    result = merge_adjacent_spans(result)
    return result


def save_report(dfs, html_file):
    total_rows = sum(len(df) for df in dfs.values())
    html_content = f"<html><head><style>table {{ width: 100%; border-collapse: collapse; }} th, td {{ border: 1px solid black; padding: 5px; text-align: left; }}</style></head><body>\n"
    html_content += f"<h2>Total issues found: {total_rows}</h2>\n"
    data_added = False  # Flag to check if any data was added

    # Loop through each DataFrame and add it to the HTML content
    for section_name, df in dfs.items():
        number_of_items = len(df)
        # Check if DataFrame is not empty
        if not df.empty:
            data_added = True  # Set the flag to True
            # Copy DataFrame to avoid modifying the original
            df_copy = df.copy()

            # Apply formatting
            df_copy[column_target] = df_copy.apply(
                lambda row: highlight_letters(row[column_target], row[column_output_name]), axis=1)

            # Convert DataFrame to HTML
            html_content += f"<h2>{section_name}, issues: {number_of_items}</h2>\n"
            html_content += df_copy.to_html(index=False, escape=False, classes='full-width') + "\n"

    # Check if no data was added
    if not data_added:
        html_content += "<p><h2>No results found</h2></p>\n"

    # End of the HTML file
    html_content += "</body></html>"

    # Write the HTML content to a file
    with open(html_file, 'w', encoding='utf-8') as file:
        file.write(html_content)


def create_report(folder_source, col_src, col_tgt, col_id, output_file='output.html', find_latin=True,
                  find_chinese=True):
    excel_data = load_excel_sheets(folder_source, col_src, col_tgt, col_id, find_latin, find_chinese)
    save_report(excel_data, output_file)
