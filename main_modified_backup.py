import zipfile
import os
import tempfile
import shutil
import re
from docx2python.utilities import replace_docx_text
import sys



def list_files(dir):
    r = []
    for root, dirs, files in os.walk(dir):
        for name in files:
            r.append(os.path.join(root, name))
    return r


def main(data_path, output_path):
    hyperlinks = []
    email_list = []

    for file in list_files(data_path):
        #print(file)
        filename = file.split("\\")[-1]
        print(filename)
        in_file_path = f"{data_path}{filename}"
        out_file_path = f"{output_path}{filename}"
        
        #unzip file
        with tempfile.TemporaryDirectory() as tmp_dir:
            # Extract copy of file footer to preserve the format
            with zipfile.ZipFile(in_file_path) as _zipfile:
                _zipfile.extractall(tmp_dir)
                
                footer_files = [f for f in list_files(f'{tmp_dir}\\word\\') if f'{tmp_dir}\\word\\footer' in f]
                os.mkdir(f"{tmp_dir}\\footer_tmp")
                for footer_file in footer_files:
                    footer_file_name = footer_file.split("\\")[-1]
                    shutil.copy(footer_file, f"{tmp_dir}\\footer_tmp\\{footer_file_name}")
            # Replace texts with spaces
            replace_docx_text(
                        in_file_path, out_file_path,
                        ("Copyright © 2014 Pacific Drilling Unpublished Work; all rights reserved.", " "),
                        ("Copyright © 2016 Pacific Drilling Unpublished Work; all rights reserved.", " "),
                        ("Copyright © 2018 Pacific Drilling Unpublished Work; all rights reserved.", " "),
                        ("Copyright © 2020 Pacific Drilling Unpublished Work; all rights reserved.", " "),
                        ("Copyright © 2010 Pacific Drilling Unpublished Work; all rights reserved.", " "),
                        ("Copyright © 2011 Pacific Drilling Unpublished Work; all rights reserved.", " "),
                        ("Copyright © 2017 Pacific Drilling Unpublished Work; all rights reserved.", " "),
                        ("Copyright © 2009 Pacific Drilling Unpublished Work; all rights reserved.", " "),
                        ("Copyright © 2012 Pacific Drilling Unpublished Work; all rights reserved.", " "),
                        ("Copyright © 2015 Pacific Drilling Unpublished Work; all rights reserved.", " "),
                        ("Copyright © 2013 Pacific Drilling Unpublished Work; all rights reserved.", " "),
                        ("Copyright © 2019 Pacific Drilling Unpublished Work; all rights reserved.", " "),
                        ("Copyright © 2019Pacific Drilling Unpublished Work; all rights reserved.", " "),
                        # ("GBS", ""),
                        html=True
            )
    
            # unzip file with modified text
            with zipfile.ZipFile(out_file_path) as _zipfile:
                _zipfile.extractall(tmp_dir)
                
                # create file path for header files and footer files
                header_files = [f for f in list_files(f'{tmp_dir}\\word\\') if f'{tmp_dir}\\word\\header' in f]
                footer_files = [f for f in list_files(f'{tmp_dir}\\word\\') if f'{tmp_dir}\\word\\footer' in f]
                
                rig_specific = [
                    "DELIVERER", "DEVELOPER", "DISCOVERER", "EXPLORER", 
                    "HIGHLANDER", "INNOVATOR", "INTEGRATOR", "INTERCEPTOR",	 
                    "INTREPID", "INVINCIBLE", "REACHER", "RESILIENT", 
                    "RESOLUTE", "RESOLVE", "VALIANT", "VENTURER", "VIKING",
                    "VOYAGER",
                    ]
                
                checklist = ["Operational Checklist", "Operasjonell Sjekkliste", "Əməliyyat Yoxlama Siyahısı", "Lista de Control Operativo", 
                    "Lista de Verificação Operacional", "Process Checklist", "Prosess-sjekkliste", "Process Yoxlama Siyahisi",
                    "Lista de Control del Proceso", "Lista de Verificação de Processo",
                    ]
                
                headerdata_text = ""
                footerdata_text = ""
                
                for footer in footer_files:
                    with open(footer, encoding="utf-8") as f:
                        footerdata = f.read()
                        footerdata_text += re.sub('<(.|\n)*?>', ' ', footerdata)
                                                
                                                                                                                        
                for header in header_files:
                    with open(header, encoding="utf-8") as f:
                        headerdata = f.read()
                        headerdata_text += re.sub('<(.|\n)*?>', ' ', headerdata)
                
                rig_specific_bool = any([word.lower() in headerdata_text.lower() for word in rig_specific])
                checklist_bool = any([word.lower() in headerdata_text.lower() for word in checklist])
                section_bool = "section_0.01" not in footerdata_text.lower()

                discard_logo = (
                    (rig_specific_bool and section_bool) or checklist_bool
                )
                for xml_file in os.listdir(f'{tmp_dir}\\word\\media\\'):
                    if xml_file == "image1.png":
                        shutil.copy(f'C:\\repos\\word\\image1.png',f'{tmp_dir}\\word\\media\\')
                    elif xml_file == "image2.png":
                        shutil.copy(f'C:\\repos\\word\\image2.png',f'{tmp_dir}\\word\\media\\')
                    else:
                        print(xml_file)
                    
            # recreate footer from copies to ensure consistent footers
            for footer_file in footer_files:
                footer_file_name = footer_file.split("\\")[-1]
                shutil.copy(f"{tmp_dir}\\footer_tmp\\{footer_file_name}", footer_file)
                      
            # recreate word document
            with zipfile.ZipFile(out_file_path, "w") as docx:
                for out_file in _zipfile.namelist():
                    docx.write(os.path.join(tmp_dir,out_file), out_file)
                    
    #Write collected hyperlinks in hyperlinks.txt           
    with open(f"{output_path}hyperlinks.txt", "w+") as f:
        for (file, hyperlink) in hyperlinks:
            f.write(file)
            f.write(", ")
            f.write(str(hyperlink))
            f.write("\n")
            
    #Write collected emails in emails.txt
    with open(f"{output_path}emails.txt", "w+") as f:
        for (file, emails) in email_list:
            f.write(file)
            f.write(", ")
            f.write(str(emails))
            f.write("\n")


if __name__ == "__main__":
    input = sys.argv[1]
    output = sys.argv[2]
    
    # input = f'C:\\A\\'
    # output= f'C:\\B\\'
    
    main(input, output)