from datetime import datetime
from pathlib import Path
import sys
import camelot.io as camelot
import fitz
# import sys
from pdf2image import convert_from_path


# RECUPERE LA DATE ET L HEURE DU JOUR
start_time = datetime.now()
print(start_time)

sys.stdout = open('ExtractionTableaux.txt', 'w', encoding='utf-8')

pdf_files = list((Path.cwd() / 'pdf').glob('*.pdf'))

for pdf in pdf_files:
    try:
        
        print('**********************************')
        print(pdf.name)
        print('**********************************')
        tables = camelot.read_pdf(f'pdf/{pdf.name}', flavor='stream', pages ='all')
            
        for table in tables:
            print(table.df)
            print('****************************************************************')
    
        print('\n\n\n')
        
    except Exception as err:
        print(err)   
        
        