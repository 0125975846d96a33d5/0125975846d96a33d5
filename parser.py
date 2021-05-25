import os
import re
import csv
import logging
from collections import defaultdict

import pdfplumber
import openpyxl
from openpyxl.styles import Alignment



pdf_dir = 'PDF'


def get_files_iter(files_dir):
    """ Generates all pdf files as scandir objects.
    -> iter
    """
    return (f for f in os.scandir(files_dir)
              if f.name.lower().endswith('.pdf'))


def pdf_to_obj(pdf_path):
    """ 
    -> pdfplumber_object
    """
    return pdfplumber.open(pdf_path)


def exists_or_create_path(folder):
    """ Assert a folder exists, if it does not
    create one.
    -> None
    """
    if not os.path.exists(folder):
        os.mkdir(folder)


def get_expected_illustr(page_text, illustr_line_idx=3):
    """ Amount of illustration expected to be in the
    pdf file.
    
    args:
        ...
        illustr_line_idx: int, line with the amount of
        images in the file.
        
    -> int
    """
    illustr_line = page_text.split('\n')[illustr_line_idx]
    return int(illustr_line.split(':')[-1].strip())


def fetch_illustr_title(text, match_limit=None):
    """ Matches illustrations codes: BL-0010, D016SA, etc.
    -> str
    """
    pattern = r'(?=[^\s]*|\s{1})[A-Z]{1,2}-{0,1}[0-9]{3,4}[A-Z]{0,2}'
    fetched_codes = [c.strip() for c in re.findall(pattern, text)
        if len(c.strip()) in [6, 7]]
    return fetched_codes[:match_limit]


def extract_image(image_path, page, image_coor=None,
                  resolution=150):
    """ Extract a image as .png file,
    with a given illustration code name.
    -> None
    """
    if image_coor is None:
        image = page.images[-1]
        image_coor = (
            image['x0'], page.height - image['y1'],
            image['x1'], page.height - image['y0'])

    image_crop = page.crop(image_coor)
    image_obj = image_crop.to_image(
        resolution=resolution)
    image_obj.save(image_path, format="PNG")


def fix_excel_corrupts_values(value):
    """ Prevent Excel from:
        - triming leading zeros from numbers,
        - converting float('nn.nn') to dates,
        - converting scale (ex. 1:100,000) to time.
    -> str
    """
    length = len(value)
    if value.isnumeric() and value.startswith('0'):
        return f'"{value}"'
    elif ((length == 5 or length == 4)
          and '.' in value and value[:2].isnumeric()):
        try:
            float(value)
            return f'"{value}"'
        except ValueError:
            pass
    elif ',' in value and ':' in value:
        return f'"{value}"'
            
    return value


def is_answer_header_line(line, ans_tokens):
    """ Check if line is an answer line,
    (A)..., (B)..., etc.
    -> str, str
    """
    for ans_section, tokens in ans_tokens.items():
        for token in tokens:
            if line.startswith(token):
                line = line.replace(token, '', 1)
                return ans_section, line.strip()
    return False


def replace_chrs(text, chrs):
    """ Multiple replaces.
    args:
        ...
        chrs = ((r1, r2), (r1, r2), ...)
    -> str
    """
    for _chr, new_chr in chrs:
        text = text.replace(_chr, new_chr)
    return text


def fix_degrees_chr(text):
    """ Replaces characters stylized as
    special ones, with actual special
    characters.
    -> str
    """
    chrs = ((' o F', ' °F'), (' oF', ' °F'))
    text = replace_chrs(text, chrs)
    degrees_match = re.findall(r'[\s-]{1}\d{1,4}o{1}',
                               text)
    if degrees_match:
        for degree_value in degrees_match:
            page_text = text.replace(
                degree_value, degree_value[:-1] + '°')
    return text


def fetch_codes_exract_images(pdf_obj, illustrations_dir):
    """ Fetches illustration codes and
    excrates images attached in the
    end of a pdf file.
    -> (illustr_codes: list, int, int)
    """
    expected_pages_illustrs = get_expected_illustr(
        pdf_obj.pages[1].extract_text())
    
    while len(pdf_obj.pages[-expected_pages_illustrs - 1].images) == 3:
        expected_pages_illustrs += 1
    
    pages_illustrs = pdf_obj.pages[-expected_pages_illustrs:]

    illustr_codes = []
    ilustrs_fetched = 0
    if expected_pages_illustrs > 0:
        for page in pages_illustrs:
            page_text = page.extract_text()

            illustr_title = fetch_illustr_title(page_text, match_limit=1)
            illustr_title = ';'.join(illustr_title)
            
            illustr_codes.append(illustr_title)
            same_codes_amount = illustr_codes.count(illustr_title)
            if same_codes_amount > 1:
                image_filename = f'{illustr_title}_{same_codes_amount}.png'
            else:
                image_filename = f'{illustr_title}.png'
            
            image_path = os.path.join(illustrations_dir, image_filename)
            if os.path.exists(image_path):
                ilustrs_fetched += 1
                continue

            if len(page.images) == 3:
                extract_image(image_path, page)
                ilustrs_fetched += 1
                
            else:
                image_coor = (0, 95, page.width, page.height - 55)
                extract_image(image_path, page, image_coor,
                              resolution=150)
                ilustrs_fetched += 1
                
    return (illustr_codes,
            expected_pages_illustrs,
            ilustrs_fetched)


def fetch_questions_answers(DATA, pdf_obj, file,
                            expected_pages_illustrs,
                            ans_tokens, illustr_codes,
                            illustrations_dir):
    """ Fetches Questions and corresponding answers
    as data rows.
    -> None
    """
    pages_idx_end = (-expected_pages_illustrs
                     if expected_pages_illustrs > 0 else None)

    fetched_questions_amount = 0 
    for idx, page in enumerate(pdf_obj.pages[:pages_idx_end],
                               start=1):
        
        page_text = page.extract_text()

        chrs = (('•', ' '), ('H2S', 'H₂S'), ('CO2', 'CO₂'))
        page_text = replace_chrs(page_text, chrs)

        page_text = fix_degrees_chr(page_text)
        
        lines = [l.strip() for l in page_text.split('\n')
                 if l.strip() and l.strip() != 'o']
        for line_idx, line in enumerate(lines):
            
            match = re.findall(r'\d{1,3}\.\s{1}', line)
            if match and line.startswith(match[0]):

                ROW = {}
                ROW['FILENAME'] = file.name
                ROW['QUESTION'] = line
                current_section = 'Q'
                for sub_line in lines[line_idx + 1:]:
                    
                    match = re.findall(r'\d{1,3}\.\s{1}', line)
                    if match and sub_line.startswith(match[0]):
                        break

                    if sub_line.startswith('o '):
                        sub_line = sub_line[1:].strip()

                    if sub_line.lower().startswith("if choice"):
                        ROW['ANS'] = sub_line.split(' ')[2]

                        ROW['QUESTION'] = ROW['QUESTION'
                                              ].strip('o').strip()
                        ROW['QUESTION'] = ROW['QUESTION'
                                              ].split(' ', 1)[-1].strip()

                        chrs = ((' -', '-'), (' - ', '-'), ('- ', '-'))
                        ROW['QUESTION'] = replace_chrs(ROW['QUESTION'], chrs)

                        for col in ['ANS A', 'ANS B', 'ANS C', 'ANS D']:
                            ROW[col] =  fix_excel_corrupts_values(ROW[col])

                        illustr_match = {
                            c for c in illustr_codes if c in ROW['QUESTION']}
                        if illustr_match:
                            ROW['ILLUSTRATION'] = ';'.join(illustr_match)
                        else:
                            ROW['ILLUSTRATION'] = None

                        # extracts images inserted within text Q\A data.
                        fetched_codes = fetch_illustr_title(
                            ROW['QUESTION'].replace('\n', ''))

                        if fetched_codes and ROW['ILLUSTRATION'] is None:
                            ROW['ILLUSTRATION'] = ';'.join(fetched_codes)

                            if page.images:
                                image_filename = f"{ROW['ILLUSTRATION']}.png"
                                image_path = os.path.join(
                                    illustrations_dir, image_filename)
                                extract_image(image_path, page)

                        fetched_questions_amount += 1
                        DATA['rows'].append(ROW)
                        break

                    answer_header_line = is_answer_header_line(
                        sub_line, ans_tokens)
                    
                    if answer_header_line:
                        current_section, sub_line = answer_header_line
                        key = f'ANS {current_section}'
                        ROW[key] = sub_line
                        continue
                        
                    if current_section == 'Q':
                        ROW['QUESTION'] += '\n' + sub_line.strip()
                    else:
                        key = f'ANS {current_section}'
                        ROW[key] += '\n' + sub_line.strip()

    return fetched_questions_amount


def parser(file, DATA, STATS, illustrations_dir,
           ans_tokens):
    """ Binds Illustrations and Question/Answers
    fetching processes.
    Updates Statistics summary data.
    -> None
    """
    pdf_obj = pdf_to_obj(file.path)

    (illustr_codes, expected_pages_illustrs,
     ilustrs_fetched) = fetch_codes_exract_images(
                        pdf_obj, illustrations_dir)

    fetched_questions_amount = fetch_questions_answers(
        DATA, pdf_obj, file, expected_pages_illustrs,
        ans_tokens, illustr_codes, illustrations_dir)

    ROW = {}
    ROW['FILENAME'] = file.name
    ROW['PAGES'] = len(pdf_obj.pages)
    ROW['EXPECTED ILLUSTR'] = expected_pages_illustrs
    ROW['FETCHED ILLUSTR'] = ilustrs_fetched
    ROW['FETCHED QUESTIONS'] = fetched_questions_amount

    STATS['rows'].append(ROW)


def csv_out(path, DATA, encoding='utf-8-sig'):
    """ Output defaultdict object as a csv file.
    args:
        ...
        DATA: defauldict(list); keys 'headers',
                                     'rows'
    -> None
    """
    with open(path, 'w', encoding=encoding) as OUT:
        OUT = csv.writer(OUT, delimiter=';',
                         escapechar='\\',
                         lineterminator='\n')
        
        OUT.writerow(DATA['headers'])
        
        for row in DATA['rows']:
            row = [row.get(h, '-') if
                   str(row.get(h, '-')).strip()
                   else '-' for h in DATA['headers']]
            OUT.writerow(row)

            
def xlsx_out(path, DATA, columns_width=False,):
    """ Outputs defaultdict object as a xlsx file.
    args:
        ...
        DATA: defauldict(list); keys 'headers',
                                     'rows'
    -> None
    """
    work_book = openpyxl.Workbook()
    work_sheet = work_book.active

    if columns_width:
        for col, width in columns_width.items():
            work_sheet.column_dimensions[
                col].width = width

    work_sheet.append(DATA['headers'])
    for row in DATA['rows']:
        row = [row[h] for h in DATA['headers']]
        work_sheet.append(row)

    # formats strings with new lines charachters
    for row in work_sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrapText=True)

    work_book.save(path)


def output(pdf_dir, DATA, STATS):
    """ Binds outputting processes.
    -> None
    """

    data_output_csv_filename = (
        f'{len(DATA["rows"])}_rows_{pdf_dir}_parsed.csv')
    csv_out(data_output_csv_filename, DATA,
            encoding='utf-8-sig')

    data_output__xlsx_filename = (
        f'{len(DATA["rows"])}_rows_{pdf_dir}_parsed.xlsx')
    columns_width = {h: 90 for h in 'BCDEF'}
    xlsx_out(data_output__xlsx_filename, DATA, columns_width)

    stats_output_filename = (
        f'{len(STATS["rows"])}_rows_{pdf_dir}_stats.csv')
    csv_out(stats_output_filename, STATS,
            encoding='utf-8-sig')
    

def main(pdf_dir, verbose=False):
    print("PDF directory:", pdf_dir)

    illustrations_dir = f'illustrations_{pdf_dir}'
    exists_or_create_path(illustrations_dir)

    STATS = defaultdict(list)
    STATS['headers'] = ['FILENAME', 'PAGES',
                        'EXPECTED ILLUSTR', 'FETCHED ILLUSTR',
                        'FETCHED QUESTIONS']

    DATA = defaultdict(list)
    DATA['headers'] = ['FILENAME','ANS', 'QUESTION',
                       'ANS A', 'ANS B', 'ANS C',
                       'ANS D', 'ILLUSTRATION']

    ans_tokens = {'A': ['(A)', '(A '], 'B': ['(B)', '(B '],
                  'C': ['(C)', '(C '], 'D': ['(D)', '(D ']}

    files = get_files_iter(pdf_dir)
    files_amount = len(os.listdir(pdf_dir))

    logging.basicConfig(
        filename='exceptions.log', level=logging.ERROR,
        format='%(asctime)s %(message)s',
        datefmt='%m/%d/%Y %H:%M:%S')

    for file_idx, file in enumerate(files):
        
        if verbose and file_idx % 10 == 0:
            print(f"{file_idx:>5} of {files_amount:>4}:",
                  file.name)
            
        try:
            parser(file, DATA, STATS,
                   illustrations_dir, ans_tokens)
        except Exception as e:
            error = e.__class__.__name__
            msg = f'{__name__}, parser: {file.name}: {error}'
            
            if verbose:
                print(f'Error: {msg}')
                
            logging.exception(msg)

    try:
        output(pdf_dir, DATA, STATS)
    except Exception as e:
        error = e.__class__.__name__
        msg = f'{__name__}, output: {file.name}: {error}'
        
        if verbose:
                print(f'Error: {msg}')
                
        logging.exception(msg)
        



if __name__ == '__main__':
    import sys

    if sys.argv[1:]:
        pdf_dir = sys.argv[1]
        
    main(pdf_dir, verbose=True)










