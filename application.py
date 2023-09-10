import os
import random
from datetime import datetime
import json
from editdocx import EditDocx
import tkinter as tk
from tkinter import filedialog, messagebox
import utils

current_date_string = datetime.now().strftime("%m/%d/%Y")

with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

def get_zip_file_path_from_user():
    zip_file_path = filedialog.askopenfilename(title="Select a ZIP file", filetypes=[("ZIP files", "*.zip")])
    return zip_file_path

class GenericWindow:
    def __init__(self, title, button_count, button_callback, extra_widgets=None):
        self.root = tk.Tk()
        self.root.title(title)
        self.center_and_resize_window()
        self.button_count = button_count
        self.button_callback = button_callback
        self.extra_widgets = extra_widgets
        self.add_buttons()
        if self.extra_widgets:
            self.extra_widgets(self.root)
        self.root.mainloop()

    def center_and_resize_window(self, width=300, height=100):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def add_buttons(self):
        for i in range(1, self.button_count + 1):
            tk.Button(self.root, text=f"Button {i}", command=lambda i=i: self.button_callback(i, self.root)).pack()

def add_label_and_buttons_to_window_a(root):
    tk.Label(root, text="Has the case been automapped?").grid(row=0, columnspan=2)
    yes_button = tk.Button(root, text="Yes", command=lambda: callback_a(root, True), width=20, height=2)
    no_button = tk.Button(root, text="No", command=lambda: callback_a(root, False), width=20, height=2)
    
    yes_button.grid(row=1, column=0)
    no_button.grid(row=1, column=1)
    
def callback_a(root, value):
    global AUTOMAP_FLAG
    AUTOMAP_FLAG = value
    print(f"Flag set to {AUTOMAP_FLAG}")
    root.destroy()
    
def add_label_and_buttons_to_window_b(root):
    tk.Label(root, text="Press continue to select .zip file.").pack()
    tk.Button(root, text="continue", command=lambda: callback_b(root), width=20, height=2).pack()
    
def callback_b(root):
    root.destroy()
    zip_file_path = get_zip_file_path_from_user()
    global zip_folder_suffix
    zip_folder_suffix = zip_file_path.split("/")[-1].split("_")[-1].split(".")[0]
    global project_subfolder_path
    project_subfolder_path = utils.populate_project_path(config["project folder path"], zip_folder_suffix, "subfolder")
    utils.extract_zip_7zip(zip_file_path, config["downloads folder path"], config["seven zip path"], config["seven zip key"])
    
def add_label_and_buttons_to_window_c(root):
    tk.Label(root, text="Take a screenshot then press continue to proceed.").pack()
    tk.Button(root, text="Take Screenshot", command=lambda: callback_c(root, "screenshot"), width=20, height=2).pack()
    tk.Button(root, text="Continue", command=lambda: callback_c(root, "continue"), width=20, height=2).pack()

def callback_c(root, action):
    global SCREENSHOT_FLAG
    if action == "screenshot":
        utils.clear_clipboard()
        utils.launch_snipping_tool()
        SCREENSHOT_FLAG = True
        print("Screenshot taken.")
    elif action == "continue":
        if SCREENSHOT_FLAG:
            print("Continued.")
            root.destroy()
        else:
            print("Please take a screenshot first.")
            
def add_label_and_buttons_to_window_d(root):
    tk.Label(root, text="Select an institution:").pack()
    options = config["institutions"]
    selected_option = tk.StringVar()
    selected_option.set(options[0])
    tk.OptionMenu(root, selected_option, *options).pack()
    tk.Button(root, text="Continue", command=lambda: callback_d(root, selected_option), width=20, height=2).pack()
    
def callback_d(root, selected_option):
    #TODO: Refactor to eliminate globals
    global SELECTED_INSTITUTION 
    SELECTED_INSTITUTION = selected_option.get()
    root.destroy()

edited_doc = EditDocx(config["template path"])
edited_doc.replace_cell_text_in_table(0,0,1,config["username"])
edited_doc.replace_cell_text_in_table(0,1,1,current_date_string)

# institution mapped?
window_a = GenericWindow("Workflow Assistant", 0, None, add_label_and_buttons_to_window_a) 
SELECTED_INSTITUTION = None
if AUTOMAP_FLAG:
    pass #TODO: Toggle yes checkbox
else:
    SCREENSHOT_FLAG = False
    # take screenshot
    window_c = GenericWindow("Workflow Assistant", 0, None, add_label_and_buttons_to_window_c) 
    image_stream = utils.get_image_from_clipboard_as_stream()
    edited_doc.insert_image(image_stream,13,width_in_inches=5)    
    #TODO: Toggle no checkbox
# select institution
window_d = GenericWindow("workflowAssistant", 0 , None, add_label_and_buttons_to_window_d) 
edited_doc.replace_cell_text_in_table(0,3,1,random.choice(config["institutions"]))
# proceed to select zip
window_b = GenericWindow("Workflow Assistant", 0, None, add_label_and_buttons_to_window_b) 
#TODO: If zip already unzipped, continue (currently freezes if unzip is attempted again)
edited_doc.replace_cell_text_in_table(0,2,1,zip_folder_suffix)

def append_to_filename(filename, string_to_append):
    name, ext = os.path.splitext(filename)
    new_name = f"{name}{string_to_append}"
    new_filename = f"{new_name}{ext}"
    return new_filename

edited_doc.save_edited_docx(os.path.join(project_subfolder_path, append_to_filename(os.path.basename(config["template path"]), f"_{zip_folder_suffix}")))

#TODO: include functionality for moving folder over to server
