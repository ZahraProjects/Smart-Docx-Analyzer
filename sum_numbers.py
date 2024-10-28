import re
from docx import Document

# Function to read text from a .docx file
def read_text_from_docx(file_path):
    # Open the .docx file and initialize an empty text string
    doc = Document(file_path)
    text = ""
    # Loop through each paragraph and add its text to the 'text' variable
    for paragraph in doc.paragraphs:
        text += paragraph.text + " "
    return text

# Function to sum numbers present in the given text
def sum_numbers_in_text(text):
    # Regular expression pattern to find integers and decimal numbers in Persian and English formats
    pattern = r"[-+]?\d+[\.,٫]?\d*"
    
    # Translation map to convert Persian digits to English
    persian_to_english = str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")
    
    # Find all numbers in the text, convert Persian numbers to English, replace Persian and English decimal separators, and calculate the sum
    total_sum = sum(
        float(num.translate(persian_to_english).replace("٫", ".").replace(",", ""))
        for num in re.findall(pattern, text)
    )
    
    return total_sum

# Specify the file path for the .docx file
file_path = "Smart-Docx-Analyzer\AI_Text_EN.docx"

try:
    # Read the text from the .docx file
    text = read_text_from_docx(file_path)
    # Calculate the sum of numbers in the text
    total = sum_numbers_in_text(text)
    # Print the result
    print("The sum of the numbers in the text:", total)
except Exception as e:
    # Print an error message if there was an issue processing the file
    print("Error processing file:", e)