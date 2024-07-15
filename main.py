from read_write_json import read_json_file, write_json_file
from read_write_docx import initialize_doc, insert_to_doc, add_header, load_doc, merge_and_save_docx, replace_text_in_docx
import math
import asyncio
import time
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename

def main():
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename(filetypes=[("json files", "*.json")]) # show an "Open" dialog box and return the path to the selected file
        
    data= read_json_file(filename)
    if not data or "dependencies" not in data:
        print("wrong file selected")
        return False
    
    base_data_dir = './data/'
    report_dir = base_data_dir + 'generated_report/'
    failed_data_dir = base_data_dir + 'failed/'
    
    doc_template_open_name = "./data/doc_template/Reporting-SAKTI-opening-template.docx"
    doc1= load_doc(doc_template_open_name)
    doc_template_close_name = "./data/doc_template/Reporting-SAKTI-closing-template.docx"
    doc2= load_doc(doc_template_close_name)
    
    ArtifactName= data["projectInfo"]["name"]
    replace_text_in_docx(doc1, "@%APP%@", ArtifactName)
    failed=[]

    start = time.time()
    
    dependencies = data["dependencies"]
    for dependency in dependencies:

        if "vulnerabilities" in dependency:
            print("- fileName = " + dependency["fileName"])
            vulnerabilities = dependency["vulnerabilities"]
            add_header(dependency["fileName"], doc1, 3)

            i = 1
            for vulnerability in vulnerabilities:
                print("- - source: " + vulnerability["source"])
                print("- - name: " + truncate_text(vulnerability["name"], 100, True))
                print("- - severity: " + vulnerability["severity"])

                if insert_to_doc(vulnerability, doc1, i, ArtifactName) == False:
                    failed.append(vulnerability)
                i+=1
    
    
    header2 = {}
    # for vulnerability in sorted_vulnerabilities:
    #     if vulnerability["Severity"] not in header2:
    #         add_header(vulnerability["Severity"], doc1, 2)
    #         header2[vulnerability["Severity"]] = True
    #     
    #     if insert_to_doc(vulnerability, doc1, i, ArtifactName) == False:
    #         failed.append(vulnerability)
    #     i+=1

    doc1.add_page_break()
    
    merge_and_save_docx(doc1, doc2, report_dir + "Depedency scan "+ArtifactName+".docx")
    
    if failed:
        print("All failed task stored in " + failed_data_dir + "Failed Append "+ArtifactName+".json")
        write_json_file(failed, failed_data_dir + "Failed Append "+ArtifactName+".json")
    
    end = time.time()
    print("Elapsed time: " + str(end-start))
    

def truncate_text(text, max_length, ellipsis=True):
    if len(text) > max_length:
        if ellipsis:
            return text[:max_length - 3] + "..."
        else:
            return text[:max_length]
    return text

if __name__ == '__main__':
    main()