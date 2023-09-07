from contextlib import redirect_stdout
from docx import Document

class EditDocx:
    
    def __init__(self, file_path):
        self.document = Document(file_path)
        self.parsed_document_data = {}
        self._parse_document()
        
    def _parse_document(self):
        self.parsed_document_data['Document'] = {'Paragraphs':{}, 'Tables':{}}
        self.parsed_document_data['Document']['Paragraphs'] = EditDocx._parse_paragraphs(self.document.paragraphs)
        self.parsed_document_data['Document']['Tables'] = EditDocx._parse_tables(self.document.tables)
        for s_num, section in enumerate(self.document.sections):
            self.parsed_document_data[f'Section_{s_num}'] = {'Header':{'Paragraphs':{}, 'Tables':{}}, 'Footer':{'Paragraphs':{}, 'Tables':{}}}
            
            if section.header:
                self.parsed_document_data[f'Section_{s_num}']['Header']['Paragraphs'] = EditDocx._parse_paragraphs(section.header.paragraphs)
                self.parsed_document_data[f'Section_{s_num}']['Header']['Tables']= EditDocx._parse_tables(section.header.tables)
            
            if section.footer:
                self.parsed_document_data[f'Section_{s_num}']['Footer']['Paragraphs'] = EditDocx._parse_paragraphs(section.footer.paragraphs)
                self.parsed_document_data[f'Section_{s_num}']['Footer']['Tables']= EditDocx._parse_tables(section.footer.tables)
    
    def print_parsed_document(self):
        EditDocx._print_dict(self.parsed_document_data)
    
    def insert_paragraph(): #will change # of paragraphs and thus affect subsequent document changes
        pass
    
    def insert_image():
        pass
    
    def replace_paragraph_text():
        
    def replace_cell_text():
        pass
    
    @staticmethod
    def _parse_tables(docx_tables):
        parsed_tables = {}
        for t_num, table in enumerate(docx_tables):
            parsed_table = EditDocx._parse_table(table)
            parsed_tables[f'Table_{t_num}'] = parsed_table
            return parsed_tables

    @staticmethod
    def _parse_table(docx_table):
        parsed_table = {}
        for r_num, row in enumerate(docx_table.rows):
            row_data = {}
            for c_num, cell in enumerate(row.cells):
                row_data[f'Col_{c_num}'] = cell.text
            parsed_table[f'Row_{r_num}'] = row_data
        return parsed_table
    
    @staticmethod
    def _parse_paragraphs(docx_paragraphs):
        parsed_paragraphs = {}
        for p_num, paragraph in enumerate(docx_paragraphs):
            parsed_paragraphs[f'Paragraph_{p_num}'] = paragraph.text
        return parsed_paragraphs
    
    @staticmethod
    def _print_dict(d, indent=0):
        for key, value in d.items():
            print('  ' * indent + str(key), end=": ")
            if isinstance(value, dict):
                print()
                EditDocx._print_dict(value, indent + 1)
            elif isinstance(value, list):
                print()
                for i, item in enumerate(value):
                    print('  ' * (indent + 1) + f"[{i}]", end=": ")
                    if isinstance(item, dict):
                        print()
                        EditDocx._print_dict(item, indent + 2)
                    else:
                        print(item)
            else:
                print(value)
                
    

if __name__ == '__main__':
    template_path = 'template.docx'
    with open(f'{template_path.split(".")[0]}_parsed.txt', 'w') as f:
        with redirect_stdout(f):
            parsed_template = EditDocx(template_path)
            parsed_template.print_parsed_document()