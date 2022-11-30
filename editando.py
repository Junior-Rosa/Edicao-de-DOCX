from datetime import datetime
from docx import Document
import pandas as pd
from docx2pdf import convert
import os



df = pd.read_excel("C:\EITI\Python\edit_doc\\nomes.xlsx")


for linha in df.index:
    documento = Document("C:\EITI\Python\edit_doc\MODELO DE CARTA.docx")
    nome = df.loc[linha, 'Nome']
    novo = f'C:\EITI\Python\edit_doc\{nome}'
    referencia = {
        'XXXX': nome,
        'DD' : str(datetime.now().day),
        'MM' : str(datetime.now().month),
        'AAAA' : str(datetime.now().year)
        
    }
    for paragrafo in documento.paragraphs:
        for cod in referencia:
            paragrafo.text = paragrafo.text.replace(cod, referencia[cod])
        #print(paragrafo.text)
    
    documento.save(novo+'.docx')
    convert(novo+'.docx')
    # os.remove(f'C:\EITI\Python\edit_doc\{nome}.docx')   
