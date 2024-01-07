import sys
import os.path
import pypdf
import re
from datetime import date
import locale
import pandas as pd

# https://regexr.com/
# https://stackoverflow.com/questions/15491894/regex-to-validate-date-formats-dd-mm-yyyy-dd-mm-yyyy-dd-mm-yyyy-dd-mmm-yyyy
re_date = r'^([0-3]*\d)\/([0-1]*\d)\/(20[1-2]\d)'

uscite = ('PRELIEVO CARTA', 'PAGAMENTO CARTA', 'PAGAMENTI DIVERSI', 'COMMISSIONE PRELIEVO EUROPA', 
    'COMMISSIONI', 'INTERESSI E COMPETENZE', 'VS.DISPOSIZIONE', 'COMMISSIONE TASSO DI CAMBIO', 
    'COMMISSIONE SERVIZIO ALERT', 'PAGAMENTO F24', 'BOLLI GOVERNATIVI')
entrate = ('ACCR. STIPENDIO-PENSIONE', 'ACCREDITO BONIFICO', 'ACCREDITO BONIFICO ESTERO', 
    'SALDO INIZIALE', 'SALDO FINALE', 'GIRO DA MIEI CONTI', 'TRASFERIMENTO IN ACCREDITO')

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
            read_pdf = pypdf.PdfReader(pdf_file)
            number_of_pages = len(read_pdf.pages)
            for p in range(0, number_of_pages):
                page = read_pdf.pages[p]
                page_content = page.extract_text()
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

            # sometimes the last line can be combined with the separator followed by the footer
            before, separator, after = line.partition('RECT_')
            try:
                op = self.parse_operation(before)
                if self.verbosity=='debug':
                    print(op)
            except OperationFormatError:
                self.append_to_last_operation(before)
            except (NoSeparatorError, NoTypeError) as e:
                raise(e)
            else:
                self.add_operation(op)
                if op[4].startswith('SALDO INIZIALE'):
                    self.balance = op[3]
                    assert op[3], 'saldo iniziale is undefined'
                elif op[4].startswith('SALDO FINALE'):
                    assert round(self.balance*100)/100 == op[3], f'final balance does not add up ({op[3]} saldo finale vs. {self.balance} calculated)'
                    self.change_state('DONE')
                else:
                    if op[3]: # entrata
                        self.balance += op[3]
                    else: # uscita
                        self.balance -= op[2]

            if separator: # end of page was reached, look for header again
                self.change_state('HEADER')

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
        for op in self.operations:
            if op[4] in ['PRELIEVO CARTA', 'PAGAMENTO CARTA', 'COMMISSIONE PRELIEVO EUROPA', 'TRASFERIMENTO IN ACCREDITO']:
                before, kw, after = op[5].partition(' PRESSO ')
                controparte.append(after)
            elif op[4] in ['ACCREDITO BONIFICO', 'ACCREDITO BONIFICO ESTERO', 'ACCR. STIPENDIO-PENSIONE']:
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
                elif op[5].startswith('ADDEBITO TELEPASS'):
                    controparte.append('TELEPASS')

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
        df = pd.DataFrame(self.operations, columns=['data operazione', 'data valuta', 'uscite', 'entrate', 'tipo', 'descrizione'])
        controparte = self.extract_controparte()
        df.insert(len(df.columns), 'controparte', controparte)
        if self.verbosity=='debug':
            print(df)
        if not filename:
            filename = self.filename.replace('.pdf', '.xlsx')
        df.to_excel(filename)
    
    def export_to_mmex(self, filename=None):
        # ID,Date,Status,Type,Account,Payee,Category,Amount,Currency,Number,Notes
        # ,DD/MM/YYYY,,Withdrawal,ACCOUNT_NAME,PAYEE_NAME,CATEGORY_NAME,-100.00,EUR,,

        export_matrix = list()

        controparte = self.extract_controparte()

        for i in range(0, len(self.operations)):
            op = self.operations[i]
            date = op[0]
            uscita = op[2]
            entrata = op[3]
            tipo = op[4]
            descrizione = op[5]
            payee = controparte[i]

            if uscita:
                type = 'Withdrawal'
                amount = -uscita
            else:
                type = 'Deposit'
                amount = entrata

            if tipo in ['SALDO INIZIALE', 'SALDO FINALE']:
                continue

            export_matrix.append((None, date, None, type, None, payee.replace(',', ''), None, None, amount, 'EUR', None, ' '.join([tipo, descrizione])))

        df = pd.DataFrame(export_matrix, columns=['ID', 'Date', 'Status', 'Type', 'Account', 'Payee', 'Category', 'SubCategory', 'Amount', 'Currency', 'Number', 'Notes'])
        
        if self.verbosity=='debug':
            print(df)
        if not filename:
            filename = self.filename.replace('.pdf', '.csv')
        df.to_csv(filename, sep=',', decimal='.', index=False)
    
def parse_file(filename):
    file_stem, file_extension = os.path.splitext(filename)
    if os.path.exists(filename) and file_extension=='.pdf':
        doc = parser(filename, verbosity='info')
        doc.parse()
        if doc.state=='DONE':
            doc.write_to_excel()
            doc.export_to_mmex()
        else:
            raise Exception(f'something didn''t go that well with {filename}')
    else:
        raise Exception(f'file {filename} does not exist, or isn''t a valid pdf file')
    
if __name__ == '__main__':
    if len(sys.argv) == 2:
        parse_file(sys.argv[1])
    else:
        print(f'provide a filename as parameter')
