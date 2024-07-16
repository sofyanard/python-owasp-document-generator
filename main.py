from read_write_json import read_json_file, write_json_file
from read_write_docx import initialize_doc, insert_to_doc, add_header, load_doc, merge_and_save_docx, replace_text_in_docx
import time
from tkinter import Tk     # from tkinter import Tk for Python 3.x
import tkinter.filedialog as fd
import concurrent.futures

def main():
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    filename=fd.askopenfilenames(title='Choose JSON File',filetypes=[("json files", "*.json")])
    start = time.time()
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        data = executor.map(get_files, filename)
    
    with concurrent.futures.ProcessPoolExecutor(max_workers=4) as executor:
        result = executor.map(generate_report, data)
    
    for r in result:
        print(r)
    
    end = time.time()
    print("Total Elapsed time: " + str(end-start))
    
def get_files(filename):
    try:
        json = read_json_file(filename)
        if not json or "dependencies" not in json:
            print("wrong file selected")

        return json
    except Exception as err:
        print(f'Error reading file: {filename}. {repr(err)}')
        # print("Error reading file: ", err)
        pass
    
def generate_report(data):
    ordered_severity = categorize_severity(data["dependencies"])

    doc_template_open_name = "./data/doc_template/Reporting-SAKTI-opening-template.docx"
    doc1= load_doc(doc_template_open_name)
    doc_template_close_name = "./data/doc_template/Reporting-SAKTI-closing-template.docx"
    doc2= load_doc(doc_template_close_name)
    
    base_data_dir = './data/'
    report_dir = base_data_dir + 'generated_report/'
    failed_data_dir = base_data_dir + 'failed/'
    ArtifactName= data["projectInfo"]["name"]
    replace_text_in_docx(doc1, "@%APP%@", ArtifactName)
    
    failed=[]
    header2=[]

    start = time.time()
    i = 1
    for severity_key, severity_data in ordered_severity.items():
        if severity_key not in header2:
            add_header(severity_key, doc1, 2)
            header2.append(severity_key)
        
        for item in severity_data:
            add_header(item["fileName"], doc1, 3)
            
            # for debugging purpose
            # print(f'- fileName = {item["fileName"]}')
            # print(f'- source = {item["source"]}')
            # print(f'- name = {truncate_text(item["name"], 100, True)}')
            # print(f'- severity = {item["severity"]}')
            
            if insert_to_doc(item, doc1, i, ArtifactName) == False:
                failed.append(item)
            i+=1  

    doc1.add_page_break()
    
    merge_and_save_docx(doc1, doc2, report_dir + "OWASP scan "+ArtifactName+".docx")
    doc_dir = report_dir + "OWASP scan "+ArtifactName+".docx"
    
    if failed:
        print("All failed task stored in " + failed_data_dir + "Failed Append "+ArtifactName+".json")
        print(f'elapsed time: {time.time()-start}')
        write_json_file(failed, failed_data_dir + "Failed Append "+ArtifactName+".json")
    else:
        return doc_dir + " is generated successfully in " + str(time.time()-start)

def categorize_severity(data):
    categorized_data = {
        "CRITICAL":[],
        "HIGH":[],
        "MEDIUM":[],
        "LOW":[],
    }
    
    # Categorize dependencies based on their severity vulnerability
    try:
        for dependency in data:
            if "vulnerabilities" in dependency:
                for severity in dependency["vulnerabilities"]:
                    if "severity" in severity:
                        severity["fileName"] = dependency["fileName"]
                        categorized_data[severity["severity"].upper()].append(severity)
        return categorized_data
    except Exception as err:
        print(f'Error in order_severity: {err}')
        print(f'Depedency data: {severity}')

def truncate_text(text, max_length, ellipsis=True):
    if len(text) > max_length:
        if ellipsis:
            return text[:max_length - 3] + "..."
        else:
            return text[:max_length]
    return text

if __name__ == '__main__':
    main()