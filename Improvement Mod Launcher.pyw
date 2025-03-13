"""
Improvement Mod Launcher
"""
from bs4 import BeautifulSoup
from configparser import ConfigParser
from ctypes import windll
import os
from pathlib import Path
from PIL import Image, ImageTk
from psutil import process_iter
import requests
from shutil import copyfileobj, rmtree
from subprocess import run, CalledProcessError, STARTUPINFO, STARTF_USESHOWWINDOW
from tempfile import mkdtemp
from threading import Thread,Event
from tkinter import filedialog, messagebox
import tkinter as tk
from winreg import OpenKey, QueryValueEx,HKEY_LOCAL_MACHINE
from zipfile import ZipFile
import customtkinter as ctk
import sys
from datetime import date
import webbrowser
import ctypes.wintypes

def debug(message:str, log:str) -> None: # Needs string message and string log message, does not return anything
    if gDebugger: # Display logs when enabled
        print(str(message), str(log)) # Print logs to console

def truncate_string(input_string, max_length=100, placeholder="..."):
    if len(input_string) > max_length:
        return input_string[:max_length - len(placeholder)] + placeholder
    return input_string

def add_log(message:str) -> None: # Adds a log message to gLogMessages and updates gLogTextBox if logging is enabled
    global gLogMessages # Required because gLogMessages is modified by append
    gLogMessages.append(message) # Modify the list of messages
    if gEnableLogs_checkbox.get() == 0: return # When logs are disabled return
    gLogTextBox.configure(state="normal") # Temporarily enable state to update text
    gLogTextBox.insert("end", message + "\n") # Insert the log message
    gLogTextBox.yview("end") # Scroll to the end of the textbox
    gLogTextBox.configure(state="disabled") # Disable user input again

def is_admin() -> bool: # Checks if the Launcher is running as admin
    try:
        debug("is_admin","Check if ran as admin " + str(bool(windll.shell32.IsUserAnAdmin())))
        return bool(windll.shell32.IsUserAnAdmin()) # Checks if app is running as admin
    except Exception as e:
        debug("is_admin","Checking failed" + str(e))
        return False

def find_label_by_name(widget) -> bool: # Finds a widget in gInterface_canvas
    for item in gInterface_canvas.find_all(): # Iterate through all canvas objects
        if gInterface_canvas.type(item) == "window": # Check if it's a window object
            if gInterface_canvas.itemcget(item, "window") == str(widget): # Check for a match
                debug("find_label_by_name","found match " + str(widget))
                return True # Found match
    return False # Did not find match

class CreateToolTip: # Displays a small popup with text after a short delay when hovering over a widget
    def __init__(self, widget, text='widget info'): # Initialize the tooltip, param widget: The widget to which the tooltip is attached, param text: The text to display in the tooltip
        debug("CreateToolTip - __init__ ","creating tip for package " + str(widget) + " " + str(text))
        self.waittime = 1000 # Time (in milliseconds) before the tooltip appears
        self.wraplength = 180 # Maximum width (in pixels) before text wraps
        self.widget = widget
        self.text = text
        self.id = None # Timer ID for scheduling tooltip appearance
        self.tw = None # Tooltip window instance
        # Bind widget events to tooltip functions
        self.widget.bind("<Enter>", self.enter) # Mouse enters widget area
        self.widget.bind("<Leave>", self.leave) # Mouse leaves widget area
        self.widget.bind("<ButtonPress>", self.leave) # Hide tooltip on button press

    def enter(self, event=None): # Called when the mouse enters the widget; schedules the tooltip display
        debug("CreateToolTip - enter ","schedule the tooltip display " + str(event))
        self.schedule()

    def leave(self, event=None): #Called when the mouse leaves the widget; cancels tooltip display
        debug("CreateToolTip - leave ","cancels the tooltip display " + str(event))
        self.unschedule()
        self.hidetip()

    def schedule(self): # Schedules the tooltip to appear after the defined wait time
        debug("CreateToolTip - schedule ","Schedules the tooltip to appear")
        self.unschedule() # Cancel any previous tooltip schedule
        self.id = self.widget.after(self.waittime, self.showtip) # Schedule tooltip

    def unschedule(self): # Cancels any scheduled tooltip display
        debug("CreateToolTip - unschedule ","Cancels any scheduled tooltip display")
        _id = self.id
        self.id = None
        if _id:
            self.widget.after_cancel(_id) # Cancel the scheduled tooltip

    def showtip(self, event=None): # Creates and displays the tooltip at the appropriate position
        # Get widget position and adjust tooltip placement
        debug("CreateToolTip - showtip ","Show the tool tip " + str(event))
        x, y, _, _ = self.widget.bbox("insert") # Get cursor position in the widget
        x += self.widget.winfo_rootx() + 25 # Adjust x position
        y += self.widget.winfo_rooty() + 38 # Adjust y position
        # Create a tooltip window
        self.tw = tk.Toplevel(self.widget)
        self.tw.wm_overrideredirect(True) # Remove window decorations
        self.tw.wm_geometry("+%d+%d" % (x, y)) # Set tooltip position
        # Create a label inside the tooltip window
        label = tk.Label(
            self.tw,
            text=self.text, # Tooltip text
            justify='left',
            background="#ffffe0", # Light yellow background
            relief='solid', # Solid border
            borderwidth=1, # Thin border
            wraplength=self.wraplength# Wrap text after specified pixels
        )
        label.pack(ipadx=1) # Add some internal padding

    def hidetip(self): # Hides the tooltip if it exists
        debug("CreateToolTip - hidetip ","Hide tool tip")
        if self.tw:
            self.tw.destroy() # Destroy the tooltip window
        self.tw = None# Reset tooltip reference

def read_write_config(fileAction:str) -> None: # r or w values only
    if fileAction not in ("r", "w"): # The operation to perform, either "r" for reading or "w" for writing
        debug("read_write_config","invalid options")
        raise ValueError("Invalid option: must be 'r' or 'w'") # Immediately raise if invalid
    try:
        if fileAction == "w": # Open file for writing
            with open(gConfigPath, 'w') as pConfigfile:
                gConfigUserData.write(pConfigfile)
                debug("read_write_config", f"Successfully wrote to {gConfigPath}")
        elif fileAction == "r":
            # Open file for reading
            gConfigUserData.read(gConfigPath)
            debug("read_write_config", f"Successfully read from {gConfigPath}")
    except (FileNotFoundError, PermissionError, IOError,Exception) as e: # Handle file access errors and give error logs
        debug("read_write_config", f"Cannot access {gConfigPath}: {e}")

def is_process_running(process_name:str) -> bool: # Check if a process with the given name is running
    for proc in process_iter(['name']):
        if proc.info['name'] == process_name:
            debug("is_process_running","Found process " + str(process_name))
            return True
    debug("is_process_running","Did not find process " + str(process_name))
    return False

def download_update_cancel() -> None: # Cancel the download
    global gCancelDownload
    gCancelDownload = True
    debug("download_update_cancel", "Download cancelled by user")

def enable_all_widgets() -> None: # Enables widgets
    for widget in gApp.winfo_children():
        if isinstance(widget, (ctk.CTkButton, ctk.CTkCheckBox, ctk.CTkComboBox,ctk.CTkOptionMenu)):
            widget.configure(state="normal") # Set widget to normal
        if isinstance(widget, ctk.CTkLabel) and widget.cget("text_color") == "#636363": # Check if the widget is a Label and color match
            widget.configure(text_color="white") # Change the text color to white
    gDownloadCancel_button.destroy()
    gDownloadCancel_label.destroy()

def disable_all_widgets() -> None: # Disable widgets
    for widget in gApp.winfo_children():
        if isinstance(widget, (ctk.CTkButton, ctk.CTkCheckBox, ctk.CTkComboBox,ctk.CTkOptionMenu)):
            widget.configure(state="disabled") # Set widget to disabled
        if isinstance(widget, ctk.CTkLabel) and widget.cget("text_color") == "white": # Check if the widget is a Label
            widget.configure(text_color="#636363") # Change the text color to gray

def download_update_thread(): # Download the new update as a separate thread
    debug("download_update_thread", "Download update as a new thread")
    # Global variables used in the function
    global gUpdate_label
    global gCancelDownload
    global gThreadStop_event
    global gStartGame_button
    global gStartGameButton_label
    global gAiDropDown
    process_names = ["age3m.exe", "age3x.exe", "age3.exe", "age3y.exe"]
    for process_name in process_names:
        if is_process_running(process_name):
            gThreadStop_event.set()
            debug("download_update_thread", "The processes is running " + str(process_name))
            add_log(f"Cannot update: {process_name} is running")
            enable_all_widgets()
            return

    zip_path: str = os.path.join(gGamePath, "mod_file.zip") # Define the zip file path
    gCancelDownload = False # Set download cancel status to False
    try: # Try downloading the update
        if os.path.exists(zip_path):
            debug("download_update_thread", "Removing mod_file.zip")
            os.remove(zip_path)
        with requests.get(gModDownloadUrl, stream=True) as response:
            response.raise_for_status() # Raise an error if the request fails
            total_size = int(response.headers.get("Content-Length", 0)) # Get file size
            block_size = 1024 # Define block size for downloading (1 KB)
            downloaded_size = 0
            with open(zip_path, "wb") as file:
                for data in response.iter_content(block_size):
                    if gCancelDownload: # Check if user Cancelled the download
                        gApp.after(0, lambda: gUpdate_label.configure(text="Download cancelled"))
                        gThreadStop_event.set()
                        debug("download_update_thread", "Download cancelled " + str(gCancelDownload))
                        enable_all_widgets()
                        break
                    if not gApp.winfo_exists(): # Check if the GUI is closed
                        gThreadStop_event.set() # kill the thread
                        enable_all_widgets()
                        return
                    file.write(data) # Write data to the file
                    downloaded_size += len(data) # Update downloaded size
                    progress_message = f"Downloaded: {downloaded_size / (1024 * 1024):.2f} MB / {total_size / (1024 * 1024):.2f} MB" # Update progress message in the UI
                    gApp.after(0, lambda: gUpdate_label.configure(text=progress_message))
        if not gCancelDownload: # If download completes successfully
            debug("download_update_thread", "Download complete")
            gApp.after(0, lambda: gUpdate_label.configure(text="Download complete. Extracting..."))
        else: # If download was Cancelled
            debug("download_update_thread", "Download cancelled")
            gApp.after(0, lambda: gUpdate_label.configure(text="Download cancelled"))
    except Exception as e: # Handle download errors
        gApp.after(0, lambda: gUpdate_label.configure(text="Download error"))
        debug("download_update_thread", "Download error:" + str(e))
        add_log("Download error:" + str(e))
        gThreadStop_event.set()
        enable_all_widgets()
        return

    if not gCancelDownload:
        try:
            with ZipFile(zip_path, "r") as zip_ref: # Open the downloaded ZIP file
                members = zip_ref.infolist()
                total_files = len(members) # Get the total number of files to extract
                for index, member in enumerate(members, start=1):
                    if gCancelDownload: # Check if user Cancelled extraction
                        debug("Extraction cancelled:", "")
                        gApp.after(0, lambda: gUpdate_label.configure(text="Extraction cancelled"))
                        enable_all_widgets()
                        return
                    target_path = os.path.join(gGamePath, member.filename) # Define extraction path
                    if member.filename.endswith("/"): # Create directories if needed
                        os.makedirs(target_path, exist_ok=True)
                    else:
                        parent_dir = os.path.dirname(target_path)
                        os.makedirs(parent_dir, exist_ok=True)
                        if os.path.exists(target_path): # Remove existing file before extraction
                            os.remove(target_path)
                        with zip_ref.open(member) as source, open(target_path, "wb") as target: # Extract and copy file contents
                            copyfileobj(source, target)
                    progress_text = f"Extracting {index}/{total_files}: {member.filename}" # Update UI with extraction progress
                    gApp.after(0, lambda text=progress_text: gUpdate_label.configure(text=text))
            debug("download_update_thread", "Updating AI...")
            read_ai_zip()
            debug("download_update_thread", "Extraction complete")
            gApp.after(0, lambda: gUpdate_label.configure(text="Extraction complete"))

            if os.path.exists(os.path.join(gGamePath, "age3m.exe")):
                gStartGame_button.configure(command = startGame_button_click)
                gStartGameButton_label.bind("<Button-1>",lambda event: gStartGame_button.invoke())
                gStartGameButton_label.configure(text_color="white")
                gStartGameButton_label.configure(cursor="hand2")
                gAiDropDown.configure(state="normal")

            if os.path.exists("D3D9.dll") or os.path.exists("D3D9.dll_Disabled"):
                gD3D9_checkbox.configure(state="normal")

        except Exception as e: # Handle extraction errors
            gApp.after(0, lambda: gUpdate_label.configure(text="Extraction error"))
            debug("download_update_thread", "Unzip error Extraction error:" + str(e))
            add_log("Cannot unzip update, update failed")
            gThreadStop_event.set()

    try: # Delete the ZIP file after extraction
        debug("download_update_thread", "Deleting the downloaded zip file...")
        os.remove(zip_path) # Remove ZIP file
        debug("download_update_thread", "zip file deleted" + str(gLast_updated))
        if not gCancelDownload: # If update was successful
            gConfigUserInfo["lastupdate"] = gLast_updated
            read_write_config("w") # Write update info to config
            read_write_config("r")
        else: # If download was Cancelled, show the update button again
            enable_all_widgets()
        gThreadStop_event.set()
        check_updates(gModDownloadUrl) # Check for future updates
        install_check()
    except Exception as e: # Handle errors while deleting ZIP file
        debug("download_update_thread", "Error deleting zip file:" + str(e))
        add_log("Error deleting temporary update zip file")
        gThreadStop_event.set()
    enable_all_widgets()

def download_update() -> None:
    global gDownloadThread
    debug("download_update", "Download update")
    gApp.after(0, lambda: gUpdate_label.configure(text="Starting download..."))
    disable_all_widgets()
    create_button("gDownloadCancel_button", "gDownloadCancel_label", "Cancel Download", 20, 40,"Cancel Download", download_update_cancel)
    gDownloadThread = Thread(target=download_update_thread, daemon=True)
    gApp.after(1, gDownloadThread.start())

def check_updates(url: str) -> str: # Function to check for mod updates by comparing the 'Last-Modified' header from the server
    global gLast_updated # Global variable to store the 'Last-Modified' header of the latest update
    global gUpdateMod_button # Global variable to reference the update button in the UI
    debug("check_updates", "checking for updates") # Log the start of the update check

    try: # Send a HEAD request to the URL to get metadata, but not the content
        response = requests.head(url, allow_redirects=True)
        response.raise_for_status() # Raise an exception if the request fails (non-2xx status code)
        gLast_updated = response.headers.get("ETag", "No Last-Modified header found") # Get the 'Last-Modified' header from the response or provide a default message if it's not found    
        update_date = response.headers.get("Last-Modified", "No Last-Modified header found")
        debug("check_updates Last updated date", gLast_updated) # Log the 'Last-Modified' header
        if gConfigUserInfo["lastupdate"] != gLast_updated: # Compare the stored 'Last-Modified' date with the value received in the response
            debug("check_updates", f"{gConfigUserInfo['lastupdate']} != {gLast_updated}") # Log the mismatch between the stored and received date
            gUpdate_label.configure(text=f"Improvement Mod New Update - {truncate_string(update_date, max_length=550, placeholder="...")}") # Update the UI to indicate that a new update is available
            gUpdateMod_button.configure(text="Update Now") # Enable the update button for the user to click
        else:  # If the dates match, the mod is up-to-date
            debug("check_updates", "UpToDate") # Log that the mod is up-to-date
            gUpdate_label.configure(text=f"Improvement Mod Up-to-date - {update_date}") # Update the UI to indicate that the mod is up-to-date
            gUpdateButton_label.configure(text="Reinstall Mod")

    except requests.RequestException as e: # Catch network-related errors during the HTTP request
        debug("check_updates Error retrieving last modified date:", str(e)) # Log the error message
        add_log("Cannot check for updates") # Add the error message to the log
        gUpdateButton_label.configure(text="Reinstall Mod")
        return "" # Return an empty string to indicate an error has occurred
    
def startGame_button_click() -> None:
    global gThreadStop_event
    debug("startGame_button_click","Starting game")
    ret = windll.shell32.ShellExecuteW(None, "runas", os.path.join(gGamePath, "age3m.exe"), "arg1", None, 1)
    if ret <= 32: # ShellExecuteW returns a value greater than 32 if successful.
        debug("startGame_button_click Failed to launch the executable with elevation",str(ret))
        add_log("Cannot start game")
    else:
        debug("startGame_button_click","Stop thread")
        gThreadStop_event.set()
        sys.exit()

def get_zip_top_level_folder_names(zip_file_path) -> set:
    """
    Open the ZIP file and return a sorted list of unique top-level folder names
    A top-level folder is defined as a folder whose name does not include a slash
    ("/") after removing any trailing slash
    """
    folder_names = set()
    try:
        with ZipFile(zip_file_path, "r") as zip_ref:
            for info in zip_ref.infolist():
                if info.is_dir(): # Check if the entry is a directory
                    folder_name = info.filename.rstrip("/") # Remove trailing slash (if present)
                    if "/" not in folder_name: # Only add if there's no further subdirectory indicated (i.e., no '/')
                        folder_names.add(folder_name)
    except Exception as e:
        debug("get_zip_top_level_folder_names Error reading " + str(zip_file_path) , str(e))
    return sorted(folder_names)

def update_file(file_text:str,file_path:str) -> None:
    try:
        with open(file_path, "w") as file: # Open file using "a" append mode 
            file.write(file_text)
            debug("update_file ImpMod_AISetting updated:", str(file_text))
    except PermissionError: # File Permission Error
        debug("update_file Permission denied", "PermissionError when writing to ImpMod_AISetting")
        add_log("Cannot update ImpMod_AISetting")
    except OSError as e: # OSError
        debug(f"update_file OS error has occurred: {e}", "OSError when writing to ImpMod_AISetting")
        add_log("Cannot update ImpMod_AISetting")
    except Exception as e: # Exception
        debug(f"update_file An unexpected error has occurred: {e}", "Exception when writing to ImpMod_AISetting")
        add_log("Cannot update ImpMod_AISetting")

def extract_ai_from_zip(folder_name: str) -> None:
    """
    Extracts all files from a specified folder inside the zip file
    and places them into the output directory, replacing existing files
    :param folder_name: The folder inside the zip to extract (relative path inside zip)
    """
    debug("extract_ai_from_zip", "Starting extraction...")
    output_dir: str = os.path.join(gGamePath, "AIM")
    zip_path: str = os.path.join(gGamePath, "ImpMod_AIs") # Ensure .zip is included if not inferred
    os.makedirs(output_dir, exist_ok=True) # Create the output directory if it doesn't exist
    try:
        with ZipFile(zip_path, 'r') as zf: # Open the zip file for reading
            for member in zf.infolist(): # Loop through all files in the zip file
                if member.filename.startswith(folder_name + '/'): # Match files within the folder
                    relative_path = os.path.relpath(member.filename, folder_name) # Compute relative path
                    target_path = os.path.join(output_dir, relative_path) # Build target file path 
                    if member.is_dir(): # If it's a directory, create it
                        os.makedirs(target_path, exist_ok=True)
                    else: # If it's a file, ensure the parent directory exists and write the file
                        os.makedirs(os.path.dirname(target_path), exist_ok=True)
                        with zf.open(member) as source, open(target_path, 'wb') as target:
                            target.write(source.read()) # Extract the file.
        debug("extract_ai_from_zip", f"Extraction complete. Files extracted to {output_dir}")
    except PermissionError as e: # Handle permission error.
        debug("extract_ai_from_zip", f"Permission denied while extracting ImpMod_AIs: {e}")
        add_log("Cannot change AI")
    except OSError as e: # Handle OS-related errors.
        debug("extract_ai_from_zip", f"OS error has occurred during extraction of ImpMod_AIs: {e}")
        add_log("Cannot change AI")
    except Exception as e: # Catch any other unexpected errors.
        debug("extract_ai_from_zip", f"An unexpected error has occurred of ImpMod_AIs: {e}")
        add_log("Cannot change AI")

def option_changed(choice) -> None:
    global gAi_label
    choice_lower = choice.lower() # Convert choice to lowercase once for reuse
    # Extract AI files only if necessary
    ai_info = ""
    if "improved_ai" in choice_lower:
        extract_ai_from_zip(choice)
        ai_info = "The default Improvement Mod AI. Built for the mod with new coding techniques, it can use advanced strategies such as kiting, build templates and more."
    elif "alistair_ai" in choice_lower:
        extract_ai_from_zip(choice)
        ai_info = "Alistair's AI is another AI built for the mod, also using new coding techniques. Overall balanced and less aggressive."
    elif "draugur_ai" in choice_lower:
        extract_ai_from_zip(choice)
        ai_info = "The famous Draugur AI, adapted for the Improvement Mod. A buffed version of the original AI in every way; does more and attacks more."
    elif "n3o_ai" in choice_lower:
        ai_info = "Murdilator's N3O AI, adapted for the Improvement Mod, built from Draugur AI; does more things overlooked in Draugur."
    elif "asian_dynasties_ai" in choice_lower:
        extract_ai_from_zip(choice)
        ai_info = "The Asian Dynasties AI, buffed and adapted for the Improvement Mod; this is also the campaign's AI."
    else:
        ai_info = f"No AI info available for {choice}"
    debug("option_changed", f"AI changed to: {choice_lower}")
    gAi_label.configure(text=ai_info) # Update the label text with AI information
    update_file(choice, os.path.join(gGamePath, "ImpMod_AISetting")) # Update the AI setting file

def read_ai_zip() -> None:
    global gSelectedAI
    global gAiDropDown
    folder_names = get_zip_top_level_folder_names(os.path.join(gGamePath, "ImpMod_AIs")) # Get folder names from the ZIP file
    # Determine the path for the button image
    if getattr(sys, 'frozen', False): # Running as a PyInstaller bundle
        button_image_path = os.path.join(sys._MEIPASS, "button.png")
    else: # Running as a script
        button_image_path = gButtonImageUrl
    try:
        button_image = Image.open(button_image_path)
    except FileNotFoundError:
        sys.exit(1)
    dropButton = ctk.CTkLabel(gApp, image=ctk.CTkImage(light_image=button_image, size=(250, 40)), text="", font=gMain_font) # Create the CTkLabel with the background image
    gInterface_canvas.create_window(20, 180, window=dropButton, anchor="w") # Add the label to the canvas
    gAiDropDown = ctk.CTkOptionMenu( # Create the OptionMenu with the same background color
        master=gApp,
        values=folder_names,
        fg_color="#600000", # Match red background color
        button_color="#800000",
        button_hover_color = "#900000",
        dropdown_fg_color="#600000",
        dropdown_text_color="white",
        corner_radius = 0,
        command=option_changed,
        width=230,
        height=30,
        dynamic_resizing=True)
    gInterface_canvas.create_window(30, 180, window=gAiDropDown,anchor="w")
    gAiDropDown.set("No AI found")
    CreateToolTip(gAiDropDown, "Selected AI will be used for both Skirmish and Multiplayer games")
    if folder_names:
        debug("read_ai_zip Found ai zip folders ", str(folder_names))
    elif os.path.exists(os.path.join(gGamePath, "ImpMod_AIs")):
        add_log("No folders found in ImpMod_AIs")
        return
    if not os.path.exists(os.path.join(gGamePath, "ImpMod_AISetting")):
        gAiDropDown.configure(state="disabled")
        add_log("Missing ImpMod_AISetting")
        return
    if not os.path.exists(os.path.join(gGamePath, "ImpMod_AIs")):
        gAiDropDown.configure(state="disabled")
        add_log("Missing ImpMod_AIs")
        return
    if not os.path.exists(os.path.join(gGamePath, "AIM")):
        gAiDropDown.configure(state="disabled")
        add_log("Missing AIM folder")
        return
    debug("read_ai_zip load ai ", str(gSelectedAI))
    try:
        with open(os.path.join(gGamePath, "ImpMod_AISetting"), "r") as file: # Open file using "a" append mode 
            gSelectedAI = file.readline().strip() # Reads the first line and removes any trailing newline
            debug("changed ai ", str(gSelectedAI))

        if gSelectedAI not in folder_names: # When there are new ai updates try and find the best selected AI candidate
            gSelectedAI = folder_names[0]

        gAiDropDown.set(gSelectedAI)
        option_changed(gSelectedAI)
    except PermissionError: # File Permission Error
        debug("read_ai_zip Permission denied", "PermissionError when reading ImpMod_AIs")
        add_log("Cannot load AI list")
    except OSError as e: # OSError
        debug(f"read_ai_zip OS error has occurred: {e}", "OSError when reading ImpMod_AIs")
        add_log("Cannot load AI list")
    except Exception as e: # Exception
        debug(f"read_ai_zip An unexpected error has occurred: {e}", "Exception when reading ImpMod_AIs")
        add_log("Cannot load AI list")  

def checkbox_D3D9_state() -> None:
    if gD3D9Var_intVar.get() == 0:
        old_name = "D3D9.dll"
        new_name = "D3D9.dll_Disabled"
    else:
        old_name = "D3D9.dll_Disabled"
        new_name = "D3D9.dll"
    try:
        if os.path.exists(old_name):
            os.rename(old_name, new_name) # Rename the file
            debug("checkbox_D3D9_state Renamed"+str(old_name)+"to", str(new_name))
        else:
            debug("checkbox_D3D9_state File not found", str(old_name))
            add_log("File not found " + str(old_name))
    except PermissionError: # File Permission Error
        debug("checkbox_D3D9_state Permission denied", "PermissionError when setting D3D9.dll")
        add_log("Permissions error when setting D3D9.dll")
    except OSError as e: # OSError
        debug(f"checkbox_D3D9_state OS error has occurred: {e}", "OSError")
        add_log("Missing D3D9.dll")
    except Exception as e: # Exception
        debug(f"checkbox_D3D9_state An unexpected error has occurred: {e}", "Exception")
        add_log("Missing D3D9.dll")

def check_files(checkFiles: list) -> bool:
    debug("check_files Checking files", str(checkFiles))
    system_paths = [r"C:\Windows\System32", r"C:\Windows\SysWOW64"]
    checkFiles_lower = [file.lower() for file in checkFiles] # Convert to lowercase
    try:
        for base_folder in system_paths:
            for root, dirs, files in os.walk(base_folder):
                files_lower = [file.lower() for file in files] # Convert to lowercase
                if any(file in files_lower for file in checkFiles_lower):
                    return True
        return False
    except PermissionError:
        debug("check_files Permission denied", "PermissionError System32 SysWOW64")
        add_log("Permissions error when reading System32 SysWOW64")
    except OSError as e:
        debug(f"check_files OS error has occurred: {e}", "OSError")
        add_log("OS error has occurred when reading System32 SysWOW64")
    except Exception as e:
        debug(f"check_files An unexpected error has occurred: {e}", "Exception")
        add_log("An unexpected error has occurred when reading System32 SysWOW64")

def enable_directplay() -> None:
    gApp.configure(state="disabled")
    try:
        # Set up the STARTUPINFO to hide the console window
        startupinfo = STARTUPINFO()
        startupinfo.dwFlags |= STARTF_USESHOWWINDOW
        result = run(
            ["dism", "/online", "/enable-feature", "/featurename:DirectPlay", "/all", "/norestart"],
            capture_output=True,
            text=True,
            check=True,
            startupinfo=startupinfo # Hide the console window
        )
        if "The operation completed successfully" in result.stdout:
            debug("enable_directplay DirectPlay has been enabled", "successfully")
            gApp.configure(state="normal")
            install_check()
        else:
            debug("enable_directplay Error", "DirectPlay might already be enabled or an error has occurred")
            add_log("Error: DirectPlay might already be enabled or an error has occurred")
            gApp.configure(state="normal")
    except CalledProcessError as e:
        debug("Error enabling DirectPlay",str(e))
        add_log("Error enabling DirectPlay")
        gApp.configure(state="normal")

def install_directx_june2010(redist_path) -> None:
    gApp.configure(state="disabled")
    temp_dir = None
    try:
        temp_dir = mkdtemp(prefix="dx_extract_") # Create a unique temporary directory for extraction
        debug(f"install_directx_june2010 Extracting {redist_path} to {temp_dir}...","")
        run([redist_path, "/Q", f"/T:{temp_dir}"],check=True) # Extract directx_Jun2010_redist.exe /Q = Silent extraction, /T:<path> = Extract to path
        debug("install_directx_june2010 Extraction complete.","")
        dxsetup_path = os.path.join(temp_dir, "DXSETUP.exe") # Run DXSETUP.exe)
        if not os.path.exists(dxsetup_path):
            add_log(f"DXSETUP.exe not found in {temp_dir}")
            gApp.configure(state="normal")
            raise FileNotFoundError(f"DXSETUP.exe not found in {temp_dir}")
        debug("install_directx_june2010 Running DXSETUP.exe in silent mode...","")
        run([dxsetup_path, "/silent"], check=True)
        debug("install_directx_june2010 DirectX9 June 2010 installation complete.","")
        install_check()
    except CalledProcessError as e:
        debug("Error installing DirectX9 June 2010:",str(e))
        add_log("Error installing DirectX9 June 2010:")
        gApp.configure(state="normal")
    except Exception as e:
        debug("Error installing DirectX9 June 2010:",str(e))
        add_log("Error installing DirectX9 June 2010:")
        gApp.configure(state="normal")
    finally: # Remove the temporary directory
        if temp_dir and os.path.exists(temp_dir):
            debug(f"Error Removing temporary directory: {temp_dir}","")
            rmtree(temp_dir, ignore_errors=True)
            install_check()

def install(fileName:str) -> None: # Install files
    gApp.configure(state="disabled")
    debug("install Installing", fileName)
    filePath = os.path.join(gGamePath, fileName) 
    debug("install Installing - filePath", filePath)
    if not os.path.exists(filePath):
        debug("install Installing - filePath", "not found")
        messagebox.showerror("Error", f"File not found: {filePath}")
        add_log("Error File not found:"+str(filePath))
        gApp.configure(state="normal")
        return
    try:
        if fileName == "msxmlenu.msi":
            run(["msiexec", "/i", filePath], check=True)
            messagebox.showinfo("Success", str(fileName) + " installed successfully!")
            debug("install installed", "successfully")
            install_check()
            return
        else:
            run([filePath], check=True)
            messagebox.showinfo("Success", str(fileName) + " installed successfully!")
            debug("install installed", "successfully")
            install_check()
            return
    except CalledProcessError as e:
        messagebox.showerror("Error", f"Installation failed: {e}")
        debug("install installed error", str(fileName))
        add_log("Install " + str(fileName) + " failed, install manually")
        gApp.configure(state="normal")
        return

def is_msxml4_installed() -> bool:
    paths = [r"SOFTWARE\Microsoft\MSXML4", r"SOFTWARE\WOW6432Node\Microsoft\MSXML4"]
    for path in paths:
        try:
            with OpenKey(HKEY_LOCAL_MACHINE, path) as key:
                version, _ = QueryValueEx(key, "Version")
                debug(f"is_msxml4_installed MSXML 4.0 Installed: Version {version}","")
                return True
        except FileNotFoundError as e:
            debug("msxml4 check error",str(e))
            add_log("msxml4 check error")
            continue  
    return False

def find_download_mirror() -> None: # Find the download link from moddb
    global gModDownloadUrl
    start_url = "https://www.moddb.com/downloads/start/286602?referer=https%3A%2F%2Fwww.moddb.com%2Fmods%2Fimprovement-mod%2Fdownloads" # Find the download page from moddb
    headers = { # Headers to mimic a real browser to get the mirror link
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://www.moddb.com/mods/improvement-mod/downloads/improvement-mod-manual-install-new"
    }
    # Fetch the HTML content
    session = requests.Session()
    try:
        response = session.get(start_url, headers=headers, timeout=5)
    except requests.exceptions.ConnectionError:
        add_log("No internet connection available")
        return
    except requests.exceptions.Timeout:
        add_log("Update check failed")
        return
    except requests.exceptions.RequestException as e:
        print(f"An error has occurred: {e}")
        add_log("Update check failed")
        return

    soup = BeautifulSoup(response.text, "html.parser") # Parse the HTML using BeautifulSoup
    # Look for the mirror link inside the page
    mirror_link = None
    for a in soup.find_all("a", href=True):
        if "/downloads/mirror/" in a["href"]:
            mirror_link = "https://www.moddb.com" + a["href"]
            break
    if mirror_link:
        gModDownloadUrl = mirror_link
        debug("find_download_mirror Mirror Link Found:",mirror_link)
    else:
        add_log("Mirror Link not Found")

def make_config() -> None: # Make a user config file to save preferences 
    debug("make_config Creating", "config")
    global gConfigPath
    global gConfigUserData
    global gConfigUserInfo
    global gGamePath
    gConfigUserData = ConfigParser() 
    gConfigPath = Path("ImprovementModLauncher.ini") # Relative path to INI file
    try:
        if gConfigPath.is_file(): # Check if the INI file exists
            read_write_config("r") # Read the config file
        else: # Set the defaults of the INI file if the file does not exists and create it
            gConfigUserData["USERINFO"] = {
            "gamepath": str(os.getcwd()),
            "lastupdate": "na",
            "showlogs": "0",
            "updatelastcheck": ""
            }
        read_write_config("w") # Wright to the config
        read_write_config("r") # Read the new config
        gConfigUserInfo = gConfigUserData["USERINFO"] # Store the INI settings information
        gGamePath = gConfigUserInfo["gamepath"] # Get the saved game path dir
        if gGamePath == "":
            gGamePath = str(os.getcwd())
        debug("make_config Config created ", "successfully")
    except PermissionError: # File Permission Error
        debug("make_config Permission denied", "PermissionError")
    except OSError as e: # OSError
        debug(f"make_config OS error has occurred: {e}", "OSError")
    except Exception as e: # Exception
        debug(f"make_config An unexpected error has occurred: {e}", "Exception")

def select_folder() -> None:
    global gGamePath
    try:
        folder_path = filedialog.askdirectory()
        if folder_path:
            folder_path = os.path.normpath(folder_path) # Automatically converts to Windows-style
            gFolderPath_label.configure(text=str(truncate_string(folder_path, max_length=180, placeholder="...")))
            gGamePath = folder_path
            gConfigUserInfo["gamepath"] = gGamePath
            read_write_config("w")
            read_write_config("r")
            debug("select_folder Change the game folder to", gGamePath)
            gInterface_canvas.pack_forget()
            main()

    except PermissionError: # File Permission Error
        debug("select_folder Permission denied", "PermissionError when setting the game folder")
        add_log("PermissionError when setting the game folder")
    except OSError as e: # OSError
        debug(f"select_folder OS error has occurred: {e}", "OSError")
        add_log("OS error has occurred when setting the game folder")
    except Exception as e: # Exception
        debug(f"select_folder An unexpected error has occurred: {e}", "Exception")
        add_log("An unexpected error has occurred when setting the game folder")

def show_logs() -> None:
    global gLogTextBox
    global gLogMessages
    global gInterface_canvas
    if find_label_by_name(gLogTextBox) and gEnableLogs_checkbox.get() == 1: return # Check if the log box is showing and it is enabled
    if gEnableLogs_checkbox.get() == 1: # Logs are enabled
        gLogTextBox = ctk.CTkTextbox(gApp, width=580, height=150,text_color="red",font= gLog_font,fg_color="#2a1107")
        gInterface_canvas.create_window(20, 640, window=gLogTextBox,anchor="w")
        gLogTextBox.configure(cursor="arrow")
        gLogTextBox.configure(state="disabled")
        for item in gLogMessages:
            gLogTextBox.configure(state="normal") # Temporarily enable state to update text
            gLogTextBox.insert("end", item + "\n") # Insert the log message
            gLogTextBox.yview("end") # Scroll to the end of the textbox
            gLogTextBox.configure(state="disabled") # Disable user input again
        gConfigUserInfo["showlogs"] = "1"
        read_write_config("w")
        read_write_config("r")
    else: # Remove the log window
        for item in gInterface_canvas.find_all(): # Iterate through all canvas objects
            if gInterface_canvas.type(item) == "window": # Check if it's a window object
                if gInterface_canvas.itemcget(item, "window") == str(gLogTextBox):
                    gInterface_canvas.delete(item)
                    break
        gConfigUserInfo["showlogs"] = "0"
        read_write_config("w")
        read_write_config("r")

def create_button(new_button,new_label,tooltip,xpos,ypos,labeltext,function_name) -> None:
    button_width = 250
    button_height = 40  
    # Determine the path for the button image
    if getattr(sys, 'frozen', False): # Running as a PyInstaller bundle
        original_image_path = os.path.join(sys._MEIPASS, "button.png")
    else: # Running as a script
        original_image_path = gButtonImageUrl
    # Open and resize the button image
    try: 
        original_image = Image.open(original_image_path)
    except FileNotFoundError:
        sys.exit(1)
    resized_image = original_image.resize((button_width, button_height)) # Resize the image
    ctk_image = ctk.CTkImage(light_image=resized_image, size=(button_width, button_height)) # Convert to CTkImage
    button_click = ctk.CTkButton( # Create the button
        master=gApp,
        text="",
        width=button_width,
        height=button_height,
        fg_color="transparent",
        font= gMain_font,
        command=function_name)
    label_image = ctk.CTkLabel( # Create an image label in front of the button
        master=gApp,
        image=ctk_image,
        text=labeltext,
        compound="center",
        text_color="white",
        font= gMain_font,
        fg_color="#2f1308",
        cursor="hand2")
    label_image.bind("<Button-1>", lambda event: button_click.invoke()) # Bind label click to trigger the button
    gInterface_canvas.create_window(xpos, ypos, window=label_image,anchor="w") # Add the button to the canvas
    CreateToolTip(label_image, str(tooltip)) # Create the tool tip
    globals()[new_button] = button_click # Update the global var button
    globals()[new_label] = label_image # Update the global var button

def interface():
    global gAi_label
    global gD3D9_checkbox
    global gD3D9Var_intVar
    global gEnableLogs_checkbox
    global gFolderPath_label
    global gInterface_canvas 
    global gLogTextBox
    global gStartGame_button
    global gStartGameButton_label
    global gUpdate_label

    gUpdate_label = ctk.CTkLabel(gApp, text="",font= gMain_font,fg_color="#2f1308",wraplength=800,anchor="w",justify='left')
    gInterface_canvas.create_window(350, 30, window=gUpdate_label,anchor='nw')

    gAi_label = ctk.CTkLabel(gApp, text="",fg_color="#2f1308",font=gAi_font,wraplength=1100,anchor="w",justify='left')
    gInterface_canvas.create_window(20, 250, window=gAi_label,anchor="nw")

    if os.path.exists("D3D9.dll"):
        gD3D9Var_intVar = ctk.IntVar(value=1)
        gD3D9_checkbox = ctk.CTkCheckBox(gApp, text="Enable DirectX12 Wrapper", variable=gD3D9Var_intVar, command=checkbox_D3D9_state,hover_color="#800000",fg_color="#600000",font= gMain_font,bg_color="#2f1308")
        gInterface_canvas.create_window(20, 320, window=gD3D9_checkbox,anchor="w")
        CreateToolTip(gD3D9_checkbox, "Enable the DirectX12 wrapper; when disabled the game will use DirectX9")
        
    elif os.path.exists("D3D9.dll_Disabled"):
        gD3D9Var_intVar = ctk.IntVar(value=0)
        gD3D9_checkbox = ctk.CTkCheckBox(gApp, text="Enable DirectX12 Wrapper", variable=gD3D9Var_intVar, command=checkbox_D3D9_state,hover_color="#800000",fg_color="#600000",font= gMain_font,bg_color="#2f1308")
        gInterface_canvas.create_window(20, 320, window=gD3D9_checkbox,anchor="w")
        CreateToolTip(gD3D9_checkbox, "Enable the DirectX12 wrapper; when disabled the game will use DirectX9")
    else:
        gD3D9_checkbox = ctk.CTkCheckBox(gApp, text="Enable DirectX12 Wrapper", variable=gD3D9Var_intVar, command=checkbox_D3D9_state,hover_color="#800000",fg_color="#600000",font= gMain_font,bg_color="#2f1308")
        gD3D9_checkbox.configure(state="disabled")
        gInterface_canvas.create_window(20, 320, window=gD3D9_checkbox,anchor="w")
        CreateToolTip(gD3D9_checkbox, "Enable the DirectX12 wrapper; when disabled the game will use DirectX9")


    gFolderPath_label = ctk.CTkLabel(gApp, text=str(truncate_string(gGamePath, max_length=180, placeholder="...")),font= gMain_font,fg_color="#2f1308",wraplength=800,anchor="w",justify='left')
    gInterface_canvas.create_window(300, 365, window=gFolderPath_label,anchor="nw")

    create_button("gUpdateMod_button", "gUpdateButton_label", "Downloads and installs new mod updates, make sure your AoE3 folder is set correctly", 20, 40,"Update Now", download_update)
    create_button("gStartGame_button", "gStartGameButton_label", "Launch age3m.exe", 20, 110,"Start Improvement Mod", startGame_button_click)
    create_button("gGamePath_button", "gGamePathButton_label", "Set the game folder path, by default the path is the root folder. Note: It is recommended to place the Improvement Mod Launcher.exe in the same folder as the game .exe", 20, 390,"Select AoE3 Folder", select_folder)
    create_button("gDirectPlay_button", "gDirectPlay_label", "Auto Enable DirectPlay", 20, 460,"Enable DirectPlay", enable_directplay)
    create_button("gDirectx_button", "gDirectx_Jun2010_label", "Install DirectX9", 300, 460,"Install DirectX9", lambda: install_directx_june2010("directx_Jun2010_redist.exe")) 
    create_button("gMsxmlenu_button", "gMsxmlenu_label", "Install MSXML4", 580, 460,"Install MSXML4", lambda: install("msxmlenu.msi")) 
    create_button("gVCredistx86_button", "gVCredistx86_label", "Install Visual C++ Redistributable ", 860, 460,"Install Visual C++", lambda: install("VC_redist.x86.exe")) 

    enableLogsVar = ctk.IntVar(value=int(gConfigUserInfo["showlogs"]))
    gEnableLogs_checkbox = ctk.CTkCheckBox(gApp, text="Show Logs", variable=enableLogsVar, command=show_logs,hover_color="#800000",fg_color="#600000",font= gMain_font,bg_color="#2f1308")
    gInterface_canvas.create_window(20, 530, window=gEnableLogs_checkbox,anchor="w")
    CreateToolTip(gEnableLogs_checkbox, "Shows logs for the launcher")

    if gEnableLogs_checkbox.get() == 1:
        gLogTextBox = ctk.CTkTextbox(gApp, width=580, height=150,text_color="red",font= gLog_font,fg_color="#2a1107")
        gInterface_canvas.create_window(20, 640, window=gLogTextBox,anchor="w")
        gLogTextBox.configure(cursor="arrow")
        gLogTextBox.configure(state="disabled")

    if is_admin() == False:
        add_log("You need to run as admin")

    if not os.path.exists(os.path.join(gGamePath, "age3m.exe")):
        gStartGame_button.unbind("<Button-1>")
        gStartGameButton_label.unbind("<Button-1>")
        gStartGameButton_label.configure(text_color="gray")
        gStartGameButton_label.configure(cursor="")
        add_log("Missing age3m.exe")

    CSIDL_PERSONAL = 5       # My Documents
    SHGFP_TYPE_CURRENT = 0   # Get current, not default value
    buf= ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
    file_path = os.path.join(buf.value, "My Games", "Age of Empires 3", "Users3", "LastProfile3.dat")    
    if not os.path.exists(file_path):
        add_log("Run age3y.exe first before running the mod")

def install_check():
    global gInterface_canvas
    global gDirectPlay_button
    global gDirectPlay_label
    global gDirectx_button
    global gDirectx_Jun2010_label
    if check_files(["dpnet.dll.mui","dpnsvr.exe.mui"]):
        gDirectPlay_button.unbind("<Button-1>")
        gDirectPlay_label.unbind("<Button-1>")
        gDirectPlay_label.configure(text_color="gray")
        gDirectPlay_label.configure(text="DirectPlay Enabled")
        gDirectPlay_label.configure(cursor="")
    else:
        debug("DirectPlay is not", "enabled")
        gDirectPlay_label.configure(text="Enable DirectPlay")

    if check_files(["d3dx9_25.dll","d3dx9_43.dll"]): # Check if DirectX June 2010 is installed
        debug("install_check DirectX9 June 2010 Redistributable is:", "Installed")
        gDirectx_button.unbind("<Button-1>")
        gDirectx_Jun2010_label.unbind("<Button-1>")
        gDirectx_Jun2010_label.configure(text_color="gray")
        gDirectx_Jun2010_label.configure(text="DirectX9 Installed")
        gDirectx_Jun2010_label.configure(cursor="")
    else:
        debug("install_check DirectX9 June 2010 Redistributable is:", "Not Installed")
        if os.path.exists("directx_Jun2010_redist.exe"):
            gDirectx_Jun2010_label.configure(text="Install Directx")
        else:
            gDirectx_button.unbind("<Button-1>")
            gDirectx_Jun2010_label.unbind("<Button-1>")
            gDirectx_Jun2010_label.configure(text_color="gray")
            gDirectx_Jun2010_label.configure(text="Directx9 Unavailable")
            gDirectx_Jun2010_label.configure(cursor="")

    if check_files(["msxml4.dll","msxml4r.dll"]):
        debug("install_check msxml4r is:", "Installed")
        gMsxmlenu_button.unbind("<Button-1>")
        gMsxmlenu_label.unbind("<Button-1>")
        gMsxmlenu_label.configure(text_color="gray")
        gMsxmlenu_label.configure(text="MSXML4 Installed")
        gMsxmlenu_label.configure(cursor="")
    else:
        debug("install_check MSXMl4 is:", "Not Installed")
        if os.path.exists("msxmlenu.msi"):
            gMsxmlenu_label.configure(text="Install MSXML4")
        else:
            gMsxmlenu_button.unbind("<Button-1>")
            gMsxmlenu_label.unbind("<Button-1>")
            gMsxmlenu_label.configure(text_color="gray")
            gMsxmlenu_label.configure(text="MSXML4 Unavailable")
            gMsxmlenu_label.configure(cursor="")

    if check_files(["msvcp140.dll", "vcruntime140.dll", "vcruntime140_1.dll"]):
        debug("install_check VC++ Redistributable is:", "Installed")
        gVCredistx86_button.unbind("<Button-1>")
        gVCredistx86_label.unbind("<Button-1>")
        gVCredistx86_label.configure(text_color="gray")
        gVCredistx86_label.configure(text="VC++ Redist Installed")
        gVCredistx86_label.configure(cursor="")
    else:
        debug("install_check MSXML4 is:", "Not Installed")
        if os.path.exists("msxmlenu.msi"):
            gVCredistx86_label.configure(text="Install VC++ Redist")
        else:
            gVCredistx86_button.unbind("<Button-1>")
            gVCredistx86_label.unbind("<Button-1>")
            gVCredistx86_label.configure(text_color="gray")
            gVCredistx86_label.configure(text="VC Redist Unavailable")
            gVCredistx86_label.configure(cursor="")
    
def main():
    debug("main Starting", "main()")
    ctk.set_appearance_mode("dark") # Set to use appearance mode for light and dark themes
    ctk.set_default_color_theme("blue") # Set color theme to blue
    debug("main Set appearance mode to system:", ctk.get_appearance_mode())
    global gInterface_canvas
    global gMain_font
    global gAi_font
    global gMain_font 
    global gAppWidth
    global gAppHeight
    global gLog_font
    gAppWidth = 1130 
    gAppHeight = 720
    gMain_font = ctk.CTkFont(family="arial", size=19,weight="bold")
    gAi_font = ctk.CTkFont(family="arial", size=15,weight="bold")
    gLog_font = ctk.CTkFont(family="arial", size=12)

    # Determine the path for the background image
    if getattr(sys, 'frozen', False): # Running as a PyInstaller bundle
        bg_image_path = os.path.join(sys._MEIPASS, "background.jpg")
    else: # Running as a script
        bg_image_path = gBackGroundImageUrl
    try: # Open the background image
        bg_image = Image.open(bg_image_path)
    except FileNotFoundError:
        sys.exit(1)

    bg_image = bg_image.resize((gAppWidth, gAppHeight)) # Resize to match window size
    bg_photo = ImageTk.PhotoImage(bg_image) # Convert to Tkinter-compatible format
    gInterface_canvas = ctk.CTkCanvas(gApp, width=gAppWidth, height=gAppHeight, highlightthickness=0,bg='black') # Create the new interface canvas
    gInterface_canvas.pack(fill="both", expand=True) # Add the canvas to the app
    gInterface_canvas.create_image(0, 0, image=bg_photo, anchor="nw") # Place the background image onto the canvas
    gApp.title("Improvement Mod Launcher - " + str(gVersion)) # Set title of the app
    gApp.resizable(0,0) # disable window resizing
    mainWindowX = (gApp.winfo_screenwidth() - gAppWidth) // 2 # Use to center the app windows along X
    mainWindowY = (gApp.winfo_screenheight() - gAppHeight) // 6 # Use to center the app windows along Y
    gApp.geometry(f"{gAppWidth}x{gAppHeight}+{mainWindowX}+{mainWindowY}") # Set the app geometry so that it is centered 
    # Run start up functions

    make_config()
    interface()
    find_download_mirror() 
    check_updates(gModDownloadUrl)
    read_ai_zip()
    install_check()

    today = date.today()
    if str(today) != gConfigUserInfo["updatelastcheck"]:
        gConfigUserInfo["updatelastcheck"] = str(today)
        read_write_config("w")
        read_write_config("r")
        try:
            GITHUB_API = "https://api.github.com"
            repo = "ageekhere/Improvement-Mod-Launcher"
            response = requests.get(f"{GITHUB_API}/repos/{repo}/releases/latest")
            latest_release = response.json()
            latest_version = latest_release.get("tag_name")
            if latest_version != gGitHubVersion and latest_version != None:
                # Create a popup window
                popup = ctk.CTkToplevel(gApp)
                popup.title("New Update")
                # Define popup dimensions
                popup_width = 300
                popup_height = 100
                # Calculate the screen's width and height
                screen_width = popup.winfo_screenwidth()
                screen_height = popup.winfo_screenheight()
                # Calculate x and y coordinates to center the window
                x = int((screen_width - popup_width) / 2)
                y = int((screen_height - popup_height) / 2)
                popup.geometry(f"{popup_width}x{popup_height}+{x}+{y}")
                 # Set the popup as a transient window to the root
                popup.transient(gApp)
                # Ensure all events are directed to the popup (modal behavior)
                popup.grab_set()
                # Focus on the popup window
                popup.focus()
                # Create a button that opens the webpage when clicked
                def open_webpage():
                    # Opens the URL in the default web browser.
                    webbrowser.open("https://github.com/ageekhere/Improvement-Mod-Launcher/releases")
                button = ctk.CTkButton(popup, text="Download at Github", command=open_webpage)
                button.pack(expand=True, padx=20, pady=20)
        except Exception as e:
            debug("Cannot Check for updates" ,str(e))
            add_log("Cannot Check for updates")

    if getattr(sys, 'frozen', False): # Close the splash screen
        pyi_splash.close()

    gApp.mainloop()

if __name__ == '__main__':
    gDebugger: bool = False # Turn on and off the debugger
    debug("Welcome to the Improvement Mod Launcher", "Creating global vars")
    ctk.deactivate_automatic_dpi_awareness() # Disable DPI scaling
    if getattr(sys, 'frozen', False): # load the splash screen 
        import pyi_splash

    gApp: ctk.CTk = ctk.CTk() # Using theme ctk
    if getattr(sys, 'frozen', False): # Running as a PyInstaller bundle so that icon.ico can be accessed
        gApp.iconbitmap(os.path.join(sys._MEIPASS, "icon.ico")) # Set the path to the icon file for bundle
    else: # Running as a script 
        gApp.iconbitmap(r"icon\icon.ico") # Set the path to the icon file for script

    gVersion:str = "1.00" # App version
    gGitHubVersion:str = "version1.00"

    gModDownloadUrl:str = "" # Hold the URL for the mirror download
    gLast_updated:str = "" # Stores the last date the file was installed
    gSelectedAI:str = "" # Currently selected AI to use
    gAiDropDown:ctk.CTkOptionMenu = None # Drop down menu for selecting an ai
    gGamePath:str = "" # stores the current game path
    gCancelDownload:bool = False # Cancel download flag
    gLogMessages:list = []
    gAppWidth = None
    gAppHeight = None
    gConfigPath:Path = None
    gConfigUserData:ConfigParser = None
    gConfigUserInfo = None
    # Interface
    gBackGroundImageUrl:str = r"interface\background.jpg"
    gButtonImageUrl:str = r"interface\button.png"
    gInterface_canvas:ctk.CTkCanvas = None
    gUpdate_label:ctk.CTkLabel = None
    gAi_label:ctk.CTkLabel = None
    gFolderPath_label:ctk.CTkLabel = None
    gDirectPlay_label:ctk.CTkLabel = None
    gDirectx_Jun2010_label:ctk.CTkLabel = None
    gMsxmlenu_label:ctk.CTkLabel = None
    gVCredistx86_label:ctk.CTkLabel = None
    gUpdateButton_label:ctk.CTkLabel = None
    gStartGameButton_label:ctk.CTkLabel = None
    gGamePathButton_label:ctk.CTkLabel = None
    gDownloadCancel_label:ctk.CTkLabel = None
    gGamePath_button:ctk.CTkButton = None
    gDirectPlay_button:ctk.CTkButton = None
    gDirectx_button:ctk.CTkButton = None
    gMsxmlenu_button:ctk.CTkButton = None
    gVCredistx86_button:ctk.CTkButton = None
    gDownloadCancel_button:ctk.CTkButton = None
    gUpdateMod_button: ctk.CTkButton = None
    gStartGame_button: ctk.CTkButton = None
    gCancelDownloadId = None
    gEnableLogs_checkbox:ctk.CTkCheckBox = None
    gD3D9_checkbox:ctk.CTkCheckBox = None
    gD3D9Var_intVar:ctk.IntVar = None
    gThreadStop_event = Event()
    gDownloadThread = None
    gMain_font:ctk.CTkFont = None
    gLog_font:ctk.CTkFont = None
    gAi_font:ctk.CTkFont = None
    gLogTextBox:ctk.CTkTextbox = None
    main()