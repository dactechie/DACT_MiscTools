import os
import docx

from collections import namedtuple
DocTup = namedtuple('doc_deets', 'uname fname')
#DeviceTup = namedtuple('dev_deets',' imei')

def get_row_text(row):
    cells_text = [ '"' + c.text.replace('\n',' ').replace('  ',' ',5).strip() +'"' for c in row.cells]
    return ",".join(cells_text)

def get_text_from_document(doct:DocTup):
    fname = 'docs/'+ doct.fname    
    doc = docx.Document(fname)
    r  = doc.tables[0].rows
    l = len(r)
    
    meta = f'"{doct.fname}","{doct.uname}",'
    #devices = [] 
    if r[l-2] and len(r[l-2].cells) > 1:
        text = meta + get_row_text(r[l-2]) + "\n"        
    if r[l-1] and len(r[l-1].cells) > 1:
        text = text + meta +  get_row_text(r[l-2]) 

    return text 

def folder_documents():
    basepath = 'docs/'
    docs =[]
    for entry in os.listdir(basepath):
        if os.path.isfile(os.path.join(basepath, entry)):
            uname = entry.replace('Mobile Device Agreement','').replace('docx','').replace('-','').replace('.',' ').strip()            
            docs.append(DocTup(uname, entry)) 
    return docs



docs  = folder_documents()
header = f'DocumentName, Username, DeviceName, IMEI_SerialNumber, Phone Number'
csv_rows = [header]

print(header)
for d in docs:
    #print(d)
    text = get_text_from_document(d)
    print(text)
    #print(f"{d['fname']},{d['uname']}, {text}")
    #csv_rows.append( f"{d['fname']},{d['uname']}, {text}")
    

#print(csv_rows)
