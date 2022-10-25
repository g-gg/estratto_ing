
# Extract movements from the balance of a ING conto arrancio (Italian branch)
This script is for those who are frustrated with the export functionality of the online banking, and still would like to collect the movements on the account in a somehow tangible way.

## Requirements
The script is written for python 3, [download](https://www.python.org/downloads/) and install on your platform is at your discretion. All required packages can be installed with [pip](https://pip.pypa.io/en/stable/) by:
`pip install -r requirements.txt`

## Usage
To run the script, run it with 
`python parse_estratto_ing.py filename.pdf`
where `filename.pdf` must point to a relative or absolute filename. The result of the operation is a `filename.xlsx` file, which can be opened conveniently in your favourite spreadsheet application.

## Description
The account balance is presented for a quarter of the year in the form of PDF document. The document is found in the customer area of [ingdirect.it](https://www.ingdirect.it/), in the section of "Conto Corrente". In the lateral menu under the group "Consulta conto" the position "Estratto conto" is found, where the documents can be downloaded.

The documents present the movements on the account in a 5 column table:
![lista movimenti](doc/lista_movimenti.png)

The first and the last row are special, the "Saldo iniziale" (opening balance), and the "Saldo finale" (final balance), representing the level of the account at the beginning and the end of the trimester respectively.

The movements in between are either active or passives, depending on the column. Unfortunately, empty columns are not resolvable in the output of [extract_text()](https://pypdf2.readthedocs.io/en/latest/modules/PageObject.html#PyPDF2._page.PageObject.extract_text). Hence the association of the amount to either actives or passives must be made on the basis of the movement description (last column). In the code there are 2 sets (uscite, entrate) to identify the respecive column.

## Testing
As I strongly advise against making any account statement public, you might test it with yours privately. Feel free to report problems as an issue in Github, and if possible, suggest potential solutions as pull requests. Please provide as much detail as possible. If the parser is initialised with `verbosity='debug'`, you may locate the problem and provide the lines that trigger the problem.






