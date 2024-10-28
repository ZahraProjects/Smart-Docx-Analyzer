# Smart-Docx-Analyzer

**Smart-Docx-Analyzer** is a Python project designed for analyzing and processing .docx files. The project includes tools for extracting numeric data, performing summation on extracted numbers, and reversing numeric sequences embedded in text. It is a flexible project that can be extended for additional features in document text and data manipulation.

## Features

1. **Numeric Summation**:
   - This script extracts all numeric values from a .docx file and calculates their sum. It supports both Persian and English numeric formats, including integers and decimal numbers.
   - **File**: sum_numbers.py

2. **Numeric Reversal**:
   - This script identifies numeric values in the text, reverses them, and saves the modified text into a new .docx file.
   - **File**: reverse_numbers.py

## Installation

To run this project, you need Python 3.x and the python-docx library. You can install the required libraries using:

bash
pip install python-docx


## Usage

### 1. Summing Numbers in a .docx File
The sum_numbers.py script reads a .docx file, extracts all numeric values, converts them if necessary, and sums them.

- **Input**: A .docx file containing Persian or English numbers.
- **Output**: The total sum of all numeric values in the text.

**Example**:  
For example, given the input text:

> "Artificial intelligence has made many advances in 2023. One of these advances is reducing errors to 2.5%."

The sum_numbers.py script would output:

> The sum of the numbers in the text: 2025.5

### 2. Reversing Numbers in a .docx File
The reverse_numbers.py script reads a .docx file, reverses any numeric sequences found in the text, and saves the modified text in a new .docx file.

- **Input**: A .docx file with numbers (backward numbers can also be included).
- **Output**: A new .docx file with reversed numeric sequences.

**Example**:  
For example, given the input text:

> "Artificial intelligence has made many advances in 2023. One of these advances is reducing errors to 2.5%."

The reverse_numbers.py script would output:

> "Artificial intelligence has made many improvements in 2012. One of these improvements is the reduction of errors to 5.2%."

## License
This project is licensed under the MIT License - see the LICENSE file for details.
