import sys
import os.path
import PyPDF2
import regex as re
from datetime import date
import locale
import pandas as pd

# https://regexr.com/
# https://stackoverflow.com/questions/15491894/regex-to-validate-date-formats-dd-mm-yyyy-dd-mm-yyyy-dd-mm-yyyy-dd-mmm-yyyy
re_date = r'^([0-3]*\d)\/([0-1]*\d)\/(20[1-2]\d)'

uscite = ('PRELIEVO CARTA', 'PAGAMENTO CARTA', 'PAGAMENTI DIVERSI', 'COMMISSIONE PRELIEVO EUROPA', 'COMMISSIONI', 'INTERESSI E COMPETENZE', 'VS.DISPOSIZIONE', 'COMMISSIONE TASSO DI CAMBIO')
entrate = ('ACCREDITO BONIFICO', 'ACCREDITO BONIFICO ESTERO', 'SALDO INIZIALE', 'SALDO FINALE')

class OperationFormatError(Exception):
    pass

class NoSeparatorError(Exception):
    pass

class NoTypeError(Exception):
    pass

class parser:
    def __init__(self, filename, verbosity='silent'):
        self.filename = filename
        self.verbosity = verbosity

    def parse(self):
        self.estimated_page = -1
        self.operations = list()
        locale.setlocale(locale.LC_NUMERIC, "it_IT")
        self.state = 'DOC_DATE' # initial state
        with open(self.filename, "rb") as pdf_file:
            read_pdf = PyPDF2.PdfFileReader(pdf_file)
            number_of_pages = read_pdf.getNumPages()
            for p in range(0, number_of_pages):
                page = read_pdf.pages[p]
                page_content = page.extractText()
                self.add_page(page_content, p, number_of_pages)

    def change_state(self, new_state):
        if self.verbosity=='debug':
            print(f'changing state from {self.state} to {new_state}')
        self.state = new_state
    
    def inc_page(self):
        self.estimated_page = self.estimated_page + 1
        if self.verbosity=='debug':
            print(f'self.estimated_page={self.estimated_page}')
    
    def parse_operation(self, line):
        m = re.match(re_date, line)
        if self.verbosity=='debug':
            print(line)

        if not m:
            raise OperationFormatError(f'no date found in {line}')
        date1 = date(int(m.groups()[2]), int(m.groups()[1]), int(m.groups()[0]))

        line = line[len(m.group()):].lstrip()
        m = re.match(re_date, line)
        if m:
            date2 = date(int(m.groups()[2]), int(m.groups()[1]), int(m.groups()[0]))
            line = line[len(m.group()):].lstrip()
        else:
            date2 = None

        euro_pos = line.find('€')
        if euro_pos < 0:
            # can happen when a date appears in the payment descriptions in the beginning of the line
            raise OperationFormatError(f'€ sign expected in {line}')

        amount_str = line[:euro_pos].strip().replace('.', '') # remove thousand separator as it confuses locale.delocalize()
        amount = locale.atof(amount_str)

        description = line[(euro_pos+1):].lstrip()

        if description.startswith('SALDO'): # INIZIALE or FINALE
            # exceptions (no separator)
            typestr = description
            description = ''
        else:
            # do the separation
            separatorstr = ' - '
            separatorpos = description.find(separatorstr)

            if separatorpos < 0:
                raise NoSeparatorError(f"separator not found in {description}")

            typestr = description[0:separatorpos].strip()
            description = description[separatorpos + len(separatorstr):]

        for e in entrate:
            if typestr == e:
                return [date1, date2, None, amount, typestr, description]

        for u in uscite:
            if typestr == u:
                return [date1, date2, amount, None, typestr, description]

        raise NoTypeError(f'{typestr} not found, expand entrate and uscite')

    def add_line(self, line, page, number_of_pages):
        if self.verbosity=='debug':
            print(line, page, number_of_pages)

        if line == 'DATA':
            return
        elif self.state == 'DOC_DATE':
            # wait for document date
            doc_date_string = 'estratto conto trimestrale al '
            if doc_date_string in line.lower():
                m = re.match(re_date, line[len(doc_date_string):])
                self.doc_date = date(int(m.groups()[2]), int(m.groups()[1]), int(m.groups()[0]))
                self.change_state('TITLE')
        elif self.state == 'TITLE':
            if 'LISTA MOVIMENTI' in line:
                self.change_state('HEADER')
        elif self.state == 'HEADER':
            if 'USCITE' in line:
                self.change_state('ROWS')
                self.inc_page()
        elif self.state == 'ROWS':
            assert self.estimated_page == page, f'actual page number ({page}) does not match expected page number ({self.estimated_page})'
            assert page <= number_of_pages, f'page ({page}) must be smaller than absolute page numbers ({number_of_pages})'

            try:
                op = self.parse_operation(line)
                if self.verbosity=='debug':
                    print(op)
            except OperationFormatError:
                # not a valid 1st line of operation

                # sometimes the last line can be combined with the separator followed by the footer
                before, separator, after = line.partition('RECT_211231')
                if (not separator) and (not after):
                    # the separator was not found
                    self.append_to_last_operation(line)
                else:
                    # the separator was found
                    if before:
                        self.append_to_last_operation(before)

                    self.change_state('HEADER')

            except (NoSeparatorError, NoTypeError) as e:
                raise(e)
            else:
                self.add_operation(op)
                if op[4].startswith('SALDO FINALE'):
                    self.change_state('DONE')


    def add_page(self, text, page, number_of_pages):
        for line in text.splitlines():
            self.add_line(line, page, number_of_pages)

    def add_operation(self, op):
        self.operations.append(op)
        if self.verbosity=='debug' or self.verbosity=='info':
            print('adding', op)

    def append_to_last_operation(self, line):
        op = self.operations[-1]
        op[5] = ' '.join([op[5], line])
        self.operations[-1] = op
        if self.verbosity=='debug' or self.verbosity=='info':
            print('appending', line)

    def extract_controparte(self):
        controparte = list()
        for op in doc.operations:
            if op[4] == 'PRELIEVO CARTA' or op[4] == 'PAGAMENTO CARTA' or op[4] == 'COMMISSIONE PRELIEVO EUROPA':
                before, kw, after = op[5].partition(' PRESSO ')
                controparte.append(after)
            elif op[4] == 'ACCREDITO BONIFICO' or op[4] == 'ACCREDITO BONIFICO ESTERO':
                m = re.match(r'(?:.*)ANAGRAFICA ORDINANTE\s(.*)\sNOTE:', op[5])
                if m:
                    controparte.append(m.groups()[0].strip())
                else:
                    raise Exception(f'could not match controparte for {op[4]} in {op[5]}')
            elif op[4] == 'PAGAMENTI DIVERSI':
                if op[5].startswith('ADDEBITO SDD'):
                    m = re.match(r'(?:.*)CREDITOR\sID\.(.*)\sID\sMANDATO', op[5])
                    if m:
                        iban_pattern = '(?:([A-Z]{2}[0-9]{2})(?=(?:[A-Z0-9]){9,30})((?:[A-Z0-9]{3,5}){2,7})([A-Z0-9]{1,3})?)'
                        test_str = m.groups()[0].strip()
                        matches = re.finditer(iban_pattern, test_str)
                        contropartestr = None

                        for matchNum, match in enumerate(matches):
                            if matchNum==0: # there shall be only one match
                                contropartestr = test_str[match.end():].strip()
                            else:
                                raise Exception(f'unexpected number of IBANS in {op[5]}')                        
                        if contropartestr:
                            controparte.append(contropartestr)
                        else:
                            raise Exception(f'could not match IBAN in {test_str}')
                    else:
                        raise Exception(f'could not match controparte for {op[4]} in {op[5]}/ADDEBITO SDD')
                elif op[5].startswith('PAGAMENTO CBILL PAGO PA'):
                    m = re.match(r'(?:.*)\sA\sFAVORE\sDI\s(.*)\sDI\sIMPORTO', op[5])
                    if m:
                        controparte.append(m.groups()[0].strip())
                    else:
                        raise Exception(f'could not match controparte for {op[4]} in {op[5]}/CBILL/PAGO PA')
                else:
                    raise Exception(f'unhandled payment {op[5]}')
            elif op[4] == 'COMMISSIONI':
                if op[5].startswith('PAGAMENTO CBILL PAGO PA'):
                    m = re.match(r'(?:.*)\sA\sFAVORE\sDI\s(.*),\sIDENTIFICATIVO\sTRANSAZIONE\s', op[5])
                    if m:
                        controparte.append(m.groups()[0].strip())
                    else:
                        raise Exception(f'could not match controparte for {op[4]} in {op[5]}/CBILL/PAGO PA')
                else:
                    raise Exception(f'unhandled commission {op[5]}')
            elif op[4] == 'VS.DISPOSIZIONE':
                m = re.match(r'(?:.*)\sA\sFAVORE\sDI\s(.*)\sBENEF.\s', op[5])
                if m:
                    controparte.append(m.groups()[0].strip())
                else:
                    raise Exception(f'could not match controparte for {op[4]} in {op[5]}')

            else:
                controparte.append(None)
        return controparte
    
    def write_to_excel(self, filename=None):
        df = pd.DataFrame(doc.operations, columns=['data operazione', 'data valuta', 'uscite', 'entrate', 'tipo', 'descrizione'])
        controparte = doc.extract_controparte()
        df.insert(len(df.columns), 'controparte', controparte)
        if self.verbosity=='debug':
            print(df)
        if not filename:
            filename = self.filename.replace('.pdf', '.xlsx')
        df.to_excel(filename)

if __name__ == '__main__':
    if len(sys.argv) == 2:
        filename, file_extension = os.path.splitext(sys.argv[1])
        if os.path.exists(sys.argv[1]) and file_extension=='.pdf':
            doc = parser(sys.argv[1], verbosity='info')
            doc.parse()
            doc.write_to_excel()
        else:
            print(f'file {sys.argv[1]} does not exist, or isn''t a valid pdf file')
    else:
        print(f'provide a filename as parameter')
