import os
import sys
import glob
from docx import Document
from docx.shared import Pt


tag = sys.argv[1]
doc = Document()

def read_all_files(folder_path):
    all_files = []
    print(f"Reading all files from: {folder_path}")

    try:
        if not os.path.exists(folder_path):
            print(f"Folder not found: {folder_path}")
            return []
        
        items = os.listdir(folder_path)
        
        for item in items:
            item_path = os.path.join(folder_path, item)
            
            if os.path.isfile(item_path):
                all_files.append(item)
                print(f"File name: {item}")
            else:
                print(f"Folder (skipped): {item}")
    
    except Exception as e:
        print(f"Error: {e}")
    
    print(f"Found {len(all_files)} files total")
    return all_files


def create_table(doc, title, records):
    table = doc.add_table(rows=2, cols=1)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = title
    hdr_cells[0].paragraphs[0].runs[0].font.size = Pt(12)
    
    for filename in records:
        data_cells = table.rows[1].cells
        data_cells[0].text = '\n'.join(records)
        data_cells[0].paragraphs[0].runs[0].font.size = Pt(10)


def combine_impact_db(combined_list):
    db_impact = clean_to_db((combined_list))
    return list(set(db_impact))


def clean_to_db(file_list):
    db = []
    for filename in file_list:
        if 'CVG' in filename:
            result = filename[filename.find('CVG'):]
        else:
            result = filename

        dot_pos = result.find('.')
        if dot_pos != -1:
            result = result[:dot_pos]
        
        db.append(result)
    
    return db

def create_document_table(doc):
    table1 = doc.add_table(rows=1, cols=2)
    table1.style = 'Table Grid'
    row1 = table1.rows[0]
    row1.cells[0].text = "Document name :"
    row1.cells[1].text = ""

    table2 = doc.add_table(rows=1, cols=4)
    table2.style = 'Table Grid'
    row2 = table2.rows[0]
    row2.cells[0].text = "Created by :"
    row2.cells[1].text = ""
    row2.cells[2].text = "Created Date :"
    row2.cells[3].text = ""

    table3 = doc.add_table(rows=1, cols=4)
    table3.style = 'Table Grid'
    row3 = table3.rows[0]
    row3.cells[0].text = "Company :"
    row3.cells[1].text = "MIMO Tech."
    row3.cells[2].text = "Department :"
    row3.cells[3].text = "BAIC"

    table4 = doc.add_table(rows=1, cols=4)
    table4.style = 'Table Grid'
    row4 = table4.rows[0]
    row4.cells[0].text = "On Production Date :"
    row4.cells[1].text = ""
    row4.cells[2].text = "Telephone :"
    row4.cells[3].text = ""

def main():
    dba_folders = glob.glob(f"./sprint/{tag}/*/DBA")
    apo_folders = glob.glob(f"./sprint/{tag}/*/APO")

    doc.add_heading("Work Instruction Template", 0)
    create_document_table(doc)

    doc.add_heading('SIR Name', level=2)
    doc.add_paragraph(f"{tag}_Enhance_CVG_Microservice_and_Fix_bug").runs[0].font.size = Pt(10)

    if not dba_folders:
        print("No DBA folders found.")
        files_dba = []
    else:
        for folder_path_dba in dba_folders:
            print(f"Found DBA folder: {folder_path_dba}")
            files_dba = read_all_files(folder_path_dba)

    if not apo_folders:
        print("No APO folders found.")
        files_apo = []
    else:
        for folder_path_apo in apo_folders:
            print(f"Found APO folder: {folder_path_apo}")
            files_apo = read_all_files(folder_path_apo)

    files_impact = combine_impact_db(files_dba + files_apo)

    if files_impact:
        doc.add_heading('Impact', level=2)
        create_table(doc, "Database", files_impact)

    if files_dba:
        doc.add_heading('Database - DBA', level=2)
        create_table(doc, "SQL Script", files_dba)

    if files_apo:
        doc.add_heading('Database - APO', level=2)
        create_table(doc, "SQL Script", files_apo)

    wi_filename = f"wi-{tag}.docx"
    doc.save(wi_filename)
    print(f"Document saved as {wi_filename}")
    

if __name__ == "__main__":
    main()
