# API Key Management Utility

## Overview

The API Key Management Utility is a Python-based tool designed to assist with managing and securing API keys in a development environment. It offers functionalities to identify potential API key leaks, storing API keys in a `.env` file, replacing API keys in code with environment variables, and identifying empty and large files in your project. 

## Features

1. **Scan Directory for API Leaks**: Identifies potential API key leaks in Python files and records them in an Excel sheet.
2. **Create `.env` File**: Extracts marked API keys from the Excel sheet and stores them in a `.env` file.
3. **Replace API Keys in Code**: Replaces marked API keys in code with environment variables or comments.
4. **Identify and Remove Empty Files**: Marks empty files in an Excel sheet, giving you an option to remove them.
5. **Identify and Remove Large Files**: Marks large files in an Excel sheet, giving you an option to remove them.
6. **Identify and Remove Empty Folders**: (Under construction)

## Prerequisites

- Python 3.x
- OpenPyXL for Excel operations (`pip install openpyxl`)
- os and re standard libraries

## How to Use

1. Run the script.
2. Follow the on-screen instructions to choose one of the available options.
3. For certain options, you will need to mark rows in an Excel file that the script generates. This is to tell the script which API keys to store in `.env` or which to replace in the source code.

### Choices

- `createExcel`: To scan your project for potential API key leaks, empty files, large files, and empty folders.
- `env` or `createEnv`: To store marked API keys from Excel into a `.env` file.
- `removeAPI`: To replace marked API keys in selected files with comments or environment variables.
- `empty`: To delete files that you've marked as empty in the Excel file.
- `largeFiles`: To delete files that you've marked as large in the Excel file.
- `emptyFolders`: Under construction.
- `exit`: To exit the program.

## File Structure

- The script operates on the current directory and allows navigation to its subdirectories.
- An Excel file named `leak_file.xlsx` is created in the current directory.
- A `.env` file is created in the directory of your choosing.

## Function Descriptions

Here are some key functions in the code:

- `on_startup()`: Initiates the user interface and provides options to manage API keys and files.
- `extract_api_keys_from_excel()`: Extracts marked API keys from an Excel file.
- `write_to_env_file()`: Writes API keys to a `.env` file.
- `update_python_files()`: Updates Python files to use environment variables for API keys.
- `find_leaks()`: Scans Python files to identify potential API key leaks.
- `find_empty_files()`: Finds empty files in the project.
- `find_large_files()`: Finds large files in the project.
- `delete_marked_files()`: Deletes marked files based on the user's choice.

## Developer Notes

- The utility uses regular expressions to identify API keys based on certain patterns.
- If you already have a `.env` file, the utility will append new API keys to it.
- The utility makes use of the OpenPyXL library to read and write Excel files for storing API key information and user choices.

## Known Issues

- The `emptyFolders` option is still under construction.

## Future Enhancements

- Option to scan non-Python files for API keys.
- Option to encrypt the `.env` file for added security.

---

Created by [Koveh.com](http://koveh.com). Feel free to contact for any issues or suggestions.
