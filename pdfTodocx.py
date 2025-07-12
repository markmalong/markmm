from pdf2docx import Converter
import os
import sys

def pdfTOword(path_input,path_output):
    for file in os.listdir(path_input): #pull files inside filepath
        cv = Converter(path_input+file) #convert using converter
        cv.convert(path_output+file.replace(".pdf","")+'.docx', start=0, end=None) #replace the ".pdf" in the file name with ".docx"
        cv.close() 
        
def main(input_path,output_path):
    pdfTOword(input_path, output_path)

if __name__ == "__main__":
    # input =f'C:\\B\\'
    # output= f'C:\\B\\'
    input = sys.argv[1]
    output = sys.argv[2]
    
    main(input,output)
    
    