import fitz
import pandas as pd
from collections import defaultdict
from openpyxl import Workbook

def extract_text_with_coordinates(pdf_path):
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"Error opening PDF file: {e}")
        return []

    page_texts = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        blocks = page.get_text("dict")["blocks"]

        page_words = []
        for block in blocks:
            if block["type"] == 0:  # block contains text
                for line in block["lines"]:
                    for span in line["spans"]:
                        for word in span["text"].split():
                            page_words.append((span["bbox"], word))
        page_texts.append(page_words)

    return page_texts

def is_table_block(bbox, page_width):
    x0, y0, x1, y1 = bbox
    block_width = x1 - x0
    block_height = y1 - y0

    # Adjust these thresholds based on typical table dimensions in your bank statements
    width_threshold = 0.3  # Adjust as needed
    height_threshold = 30  # Adjust as needed

    # Heuristic to determine if a block is part of a table
    if (x1 - x0) / page_width > width_threshold and block_height < height_threshold:
        return True
    else:
        return False

def separate_text(text):
  # Separate text into different columns based on conditions
  new_words = []
  current_word = ''
  consecutive_spaces = 0

  for char in text:
    if char == ' ':
      consecutive_spaces += 1
      if consecutive_spaces >= 5:  # Check for at least 5 spaces
        if current_word:
          new_words.append(current_word.strip())
        current_word = ''
        consecutive_spaces = 0
      else:
        if current_word and current_word[-1] == ' ':
          if current_word.split()[-1].isdigit() and char.isdigit():
            new_words.append(current_word.strip())
            current_word = char
          else:
            new_words.append(current_word.strip())
            current_word = char
        else:
          current_word += char
    else:
      consecutive_spaces = 0
      current_word += char

  if current_word:
    new_words.append(current_word.strip())

  return new_words


def process_words(page_words, page_width):
    table_lines = defaultdict(list)
    last_y = None

    for word in page_words:
        if len(word) == 2 and isinstance(word[0], tuple):  # Check if word is a tuple containing bbox and text
            bbox, text = word
            x0, y0, x1, y1 = bbox

            if is_table_block(bbox, page_width):
                if last_y is None or abs(last_y - y0) < 10:  # Adjust the threshold as needed
                    table_lines[y0].append((x0, text))
                else:
                    yield table_lines
                    table_lines = defaultdict(list)
                    table_lines[y0].append((x0, text))

                last_y = y0

    if table_lines:
        yield table_lines

def process_tables(page_texts, doc):
    all_tables = []

    for page_num, page_words in enumerate(page_texts):
        page_width = doc.load_page(page_num).rect.width
        for table_lines in process_words(page_words, page_width):
            table_data = defaultdict(list)
            for y0, line in table_lines.items():
                line = sorted(line)  # Sort based on x0 (first element in the tuple)
                combined_line = []
                for x0, text in line:
                    separated_text = separate_text(text)
                    for item in separated_text:
                        combined_line.append((item, x0))

                for i, (text, x0) in enumerate(combined_line):
                    table_data[i].append(text)
            df = pd.DataFrame.from_dict(table_data, orient='index').transpose()
            all_tables.append(df)

    return pd.concat(all_tables, ignore_index=True)

def save_table_to_excel(df, output_path):
    wb = Workbook()
    ws = wb.active

    for r in df.to_records(index=False):
        ws.append(list(r))

    try:
        wb.save(output_path)
    except Exception as e:
        print(f"Error saving Excel file: {e}")

def main(pdf_path, extracted_excel_path):
    page_texts = extract_text_with_coordinates(pdf_path)
    if not page_texts:
        print("No text extracted from PDF. Exiting.")
        return
    
    doc = fitz.open(pdf_path)
    concatenated_df = process_tables(page_texts, doc)
    save_table_to_excel(concatenated_df, extracted_excel_path)

if __name__ == "__main__":
    pdf_path = "test9.pdf"
    extracted_excel_path = "extracted_tables.xlsx"

    main(pdf_path, extracted_excel_path)
