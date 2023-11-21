![image](https://github.com/wtigga/lockitcheck/assets/7037184/fabadfb7-41f7-4ad4-bc87-a6cd7d0ee78e)

LockitCheckM is a Python script with GUI designed to identify characters in target language of a bilingual localization file that are not supposed to be there. It was designed to find Chinese characters (including non-European punctuation marks) and Latin letters (for non-latin alphabetic languages, like Russian), with a mechanism to easily define exclusions.

## Features
- **Offline**: No data transfer outside your computer.
- **Bulk Processing**: Easily process multiple XLSX files by selecting a folder.
- **Custom Column Selection**: Select columns for source language, target language, and string ID through intuitive dropdown menus.
- **Target Language Checks**: Includes options to check for Latin characters or Chinese (CJK) characters in the target language column, useful for ensuring language integrity.
- **Exclusion Lists**: Customize your checks by excluding specific words or using regular expressions. Simply add them to `exclude_elements.txt` and `exclude_regext.txt`. The current build of the program already have exclusion rules for Genshin Impact localization.

## How to Use
1. **Launch** lockitcheck.exe
2. **Browse Folder**: Click 'Browse Folder' and select the  /folder/  containing your bilingual XLSX files.
3. **Select Columns**: Choose the appropriate columns for the Source, Target, and ID from the dropdown menus.
4. **Check Options**: Enable checks for Latin or Chinese characters in the target language, as needed.
5. **Process**: Click 'Process' to run the script and generate an HTML report of the findings.

Once finished, a browser with the report file will pop up.

## Setting Exclusions

To tailor the checks to your specific needs, you can set exclusions:

- **Word Exclusions**: Add words to `exclude_elements.txt`. Each word should be on a new line.
- **Regex Exclusions**: For more complex patterns, add regular expressions to `exclude_regext.txt`, with one expression per line.