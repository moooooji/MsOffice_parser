import os
import zipfile
import xml.etree.ElementTree as ET
import olefile

def distinguish_file_type(file_path):
    with open(file_path, 'rb') as file:
        
        magic_num = file.read(4)
    
    ooxml_magic_num = b'\x50\x4b\x03\x04'
    
    if magic_num == ooxml_magic_num:
        return "OOXML 파일"
        
    else:
        return "OLE 파일"

def compress_ooxml_to_zip(input_file_path):
    
    directory, file_name = os.path.split(input_file_path)
    
    full_length = len(file_name)
    
    new_file_name = file_name[0:full_length-5] + ".zip"
    
    new_file_path = os.path.join(directory, new_file_name)
    
    os.rename(input_file_path, new_file_path)
    
    return new_file_path

def decompress_ooxml_to_zip(zip_file_path, output_directory):
    try:
        with zipfile.ZipFile(zip_file_path, 'r') as zip_file:
            zip_file.extractall(output_directory)
            xml_files = [file for file in os.listdir(output_directory) if file.endswith('.xml')]
            
            return xml_files
        
    except Exception as e:
        print(f"압축 해제 오류: {e}")
        
def analyze_OLE_files(file_path):
    
    OLE = olefile.OleFileIO(file_path)
    
    objects = OLE.listdir()
    
    for obj in objects:
        
        if obj == ["WordDocument"]:
            print("doc 파일입니다.")
            return
        
        if obj == ["PowerPoint Document"]:
            print("ppt 파일입니다.")
            return
        
        if obj == ["Workbook"]:
            print("xls 파일입니다.")
            return
    
if __name__ == "__main__":
    file_path = input("파일 경로를 입력하세요 : ")

    if distinguish_file_type(file_path) == "OOXML 파일":
        
        new_file_path = compress_ooxml_to_zip(file_path)
            
        output_directory = input("압축 해제할 디렉토리 경로를 입력하세요: ")
        
        xml_files = decompress_ooxml_to_zip(new_file_path, output_directory)
        
        
        for xml_file in xml_files:
            
            xml_file_path = os.path.join(output_directory, xml_file)
            
            tree = ET.parse(xml_file_path)
        
            root = tree.getroot()
            
            ns_URI = root.tag.split('}')[0][1:]

            namespace = {'ns': ns_URI}

            overrides = root.findall(".//ns:Override", namespaces=namespace)

            for override in overrides:
                part_name = override.get('PartName')

                if part_name == '/xl/workbook.xml':
                    print("xlsx 파일입니다.")
                    

                if part_name == '/ppt/slides/slide1.xml':
                    print("pptx 파일입니다.")
                    

                if part_name == '/word/document.xml':
                    print("docx 파일입니다.")
                
    else:
        analyze_OLE_files(file_path)