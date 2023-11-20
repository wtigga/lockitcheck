# LockitCheck

LockitCheck is a Python script with GUI designed to identify issues in bilingual localization files. It's an essential tool for anyone working with multilingual datasets, particularly useful in the context of software localization, where accuracy and consistency across languages are crucial.

## Features

- **Bulk Processing**: Easily process multiple XLSX files by selecting a folder.
- **Custom Column Selection**: Select columns for source language, target language, and string ID through intuitive dropdown menus.
- **Target Language Checks**: Includes options to check for Latin characters or Chinese (CJK) characters in the target language column, useful for ensuring language integrity.
- **Exclusion Lists**: Customize your checks by excluding specific words or using regular expressions. Simply add them to `exclude_elements.txt` and `exclude_regext.txt`.

## How to Use

1. **Browse Folder**: Click 'Browse Folder' and select the folder containing your bilingual XLSX files.
2. **Select Columns**: Choose the appropriate columns for the Source, Target, and ID from the dropdown menus.
3. **Check Options**: Enable checks for Latin or Chinese characters in the target language, as needed.
4. **Process**: Click 'Process' to run the script and generate an HTML report of the findings.

## Setting Exclusions

To tailor the checks to your specific needs, you can set exclusions:

- **Word Exclusions**: Add words to `exclude_elements.txt`. Each word should be on a new line.
- **Regex Exclusions**: For more complex patterns, add regular expressions to `exclude_regext.txt`, with one expression per line.