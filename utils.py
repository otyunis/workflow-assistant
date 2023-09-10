import subprocess
import os
from PIL import ImageGrab
from io import BytesIO
import pyperclip

def launch_snipping_tool():
    try:
        subprocess.Popen("SnippingTool.exe")
    except Exception as e:
        return f"Failed to launch Snipping Tool: {e}"

def get_image_from_clipboard_as_stream():
    image = ImageGrab.grabclipboard()
    image_byte_arr = BytesIO()
    image.save(image_byte_arr, format='PNG')
    image_stream = BytesIO(image_byte_arr.getvalue())
    return image_stream

def clear_clipboard():
    pyperclip.copy('')

def is_clipboard_empty():
    return '' == pyperclip.paste()

def get_clipboard():
    return pyperclip.paste()
    
def extract_zip_7zip(zip_file_path, output_folder, seven_zip_path, seven_zip_key):

    # Check if 7-Zip is installed
    # try:
    #     subprocess.run(["7z"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    # except FileNotFoundError:
    #     print("7-Zip is not installed or not added to the system path.") #TODO: Address 7-Zip installation/adding to system path
    #     return

    # Validate file path and output folder
    if not os.path.exists(zip_file_path):
        print(f"The zip file {zip_file_path} does not exist.")
        return
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Construct the 7-Zip command
    command = [seven_zip_path, "x", f"-o{output_folder}", zip_file_path]
    if seven_zip_key:
        command.append(f"-p{seven_zip_key}")

    # Execute the command
    result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    if result.returncode == 0:
        print("Extraction successful.")
    else:
        print("An error occurred during extraction.")
        
def populate_project_path(project_folder_path, project_name, project_subfolder_name): 
#TODO: Improve project path naming (e.g., distinguish b/w project parent folder and project folder)
#TODO: make project structure generic
    
    if not os.path.exists(project_folder_path):
        os.mkdir(project_folder_path)
    
    project_path = os.path.join(project_folder_path,project_name)
    if not os.path.exists(project_path):
        os.mkdir(project_path)
        
    project_subfolder_path = os.path.join(project_path,project_subfolder_name)
    if not os.path.exists(project_subfolder_path):
        os.mkdir(project_subfolder_path)
    return project_subfolder_path
