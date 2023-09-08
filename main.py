import os
from datetime import datetime
import json
from editdocx import EditDocx
import utils

current_date_string = datetime.now().strftime("%m/%d/%Y")

with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

username = config["username"]
institutions = config["institutions"]
template_path = config["template path"]
project_folder = config["project folder path"]
downloads_folder = config["downloads folder path"]
seven_zip_key = config["seven zip key"]
seven_zip_path = config["seven zip path"]

zip_file_path = utils.get_zip_file_path_from_user()
zip_folder_suffix = zip_file_path.split("/")[-1].split("_")[-1].split(".")[0]
# utils.extract_zip_7zip(zip_file_path, downloads_folder, seven_zip_path, seven_zip_key)


# edited_doc = EditDocx(template_path)
# edited_doc.replace_cell_text_in_table(0,0,1,username)
# edited_doc.replace_cell_text_in_table(0,1,1,current_date_string)
# edited_doc.replace_cell_text_in_table(0,2,1,random_string)
# edited_doc.replace_cell_text_in_table(0,3,1,random.choice(institutions))
# image_stream = utils.get_image_from_clipboard_as_stream()
# edited_doc.insert_image(image_stream,13,width_in_inches=5)
# edited_doc.save_edited_docx("template_edited.docx")