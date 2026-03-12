import os
import platform
import subprocess
from datetime import datetime
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# 导入原有的main函数
from main import generate as generate_calendar


class ExamCalendarApp:
    def __init__(self):
        self.root = ttk.Window(title="CIE Exam Calendar", themename="litera")
        self.root.geometry("500x400")
        self.root.minsize(450, 350)

        self.setup_ui()
        self.center_window()

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def setup_ui(self):
        # Main container with padding
        main = ttk.Frame(self.root, padding="20")
        main.pack(fill=BOTH, expand=YES)

        # Title
        ttk.Label(main, text="CIE Exam Calendar", font=("", 16, "bold")).pack(anchor=W, pady=(0, 20))

        # Student name
        ttk.Label(main, text="Student Name", font=("", 10)).pack(anchor=W)
        self.name_var = ttk.StringVar(value="Tom")
        self.name_entry = ttk.Entry(main, textvariable=self.name_var, font=("", 10))
        self.name_entry.pack(fill=X, pady=(0, 15))

        # Subject codes label
        ttk.Label(main, text="Subject Codes (one per line)", font=("", 10)).pack(anchor=W)

        # Text area for subject codes
        text_frame = ttk.Frame(main)
        text_frame.pack(fill=BOTH, expand=YES, pady=(0, 10))

        self.codes_text = ttk.Text(text_frame, height=8, font=("", 10), wrap=NONE)
        self.codes_text.pack(side=LEFT, fill=BOTH, expand=YES)

        # Scrollbar for text area
        scrollbar = ttk.Scrollbar(text_frame, orient=VERTICAL, command=self.codes_text.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.codes_text.configure(yscrollcommand=scrollbar.set)

        # Default values
        default_codes = """9231/14
9231/44
9700/14
9700/24
9700/38"""
        self.codes_text.insert("1.0", default_codes)

        # Button frame
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=X, pady=10)

        # Generate button
        self.generate_btn = ttk.Button(
            btn_frame,
            text="Generate",
            command=self.generate_calendar,
            bootstyle=SUCCESS,
            width=15
        )
        self.generate_btn.pack(side=LEFT, padx=(0, 5))

        # Clear button
        self.clear_btn = ttk.Button(
            btn_frame,
            text="Clear",
            command=self.clear_codes,
            bootstyle=SECONDARY,
            width=10
        )
        self.clear_btn.pack(side=LEFT)

        # Status frame
        status_frame = ttk.Frame(main)
        status_frame.pack(fill=X, pady=(10, 5))

        ttk.Label(status_frame, text="Status:", font=("", 10)).pack(side=LEFT)
        self.status_var = ttk.StringVar(value="Ready")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, font=("", 10))
        self.status_label.pack(side=LEFT, padx=(5, 0))

        # Output file frame
        output_frame = ttk.Frame(main)
        output_frame.pack(fill=X)

        self.output_var = ttk.StringVar(value="")
        self.output_label = ttk.Label(output_frame, textvariable=self.output_var, font=("", 9))
        self.output_label.pack(side=LEFT, fill=X, expand=YES)

        # Open folder button (hidden initially)
        self.open_btn = ttk.Button(
            output_frame,
            text="📁 Open",
            command=self.open_output_folder,
            bootstyle=INFO,
            state=DISABLED
        )

    def get_subject_codes(self):
        """Get subject codes from text area"""
        codes_text = self.codes_text.get("1.0", END).strip()
        if not codes_text:
            return []

        # Split by lines and filter empty lines
        codes = [code.strip() for code in codes_text.split("\n") if code.strip()]
        return codes

    def generate_calendar(self):
        """Generate calendar file"""
        # Get inputs
        student_name = self.name_var.get().strip()
        if not student_name:
            messagebox.showerror("Error", "Please enter student name")
            return

        subject_codes = self.get_subject_codes()

        # Update UI
        self.generate_btn.configure(state=DISABLED, text="Generating...")
        self.status_var.set("Generating calendar...")
        self.root.update()

        try:
            # Call original main function
            generate_calendar(student_name, subject_codes)

            # Find the latest generated file
            timestamp = datetime.now().timestamp()
            expected_filename = f"{student_name}-Exams2026S_{int(timestamp)}.ics"

            # Check if file exists (there might be slight timestamp difference)
            files = [f for f in os.listdir(".") if f.startswith(f"{student_name}-Exams2026S_") and f.endswith(".ics")]
            if files:
                # Get the most recent file
                latest_file = max(files, key=lambda f: os.path.getctime(f))
                self.output_var.set(f"✓ {latest_file}")
                self.status_var.set("Success!")

                # Show and enable open folder button
                self.open_btn.pack(side=RIGHT)
                self.open_btn.configure(state=NORMAL)
                self.current_file = latest_file
            else:
                self.output_var.set("⚠ File generated but not found")
                self.status_var.set("Warning")

        except Exception as e:
            self.status_var.set("Error!")
            self.output_var.set(f"✗ {str(e)[:50]}...")
            messagebox.showerror("Error", f"Failed to generate calendar:\n{str(e)}")
        finally:
            self.generate_btn.configure(state=NORMAL, text="Generate")

    def clear_codes(self):
        """Clear subject codes text area"""
        self.codes_text.delete("1.0", END)
        self.status_var.set("Cleared")
        self.output_var.set("")

    def open_output_folder(self):
        """Open folder containing the generated file"""
        if hasattr(self, 'current_file'):
            folder_path = os.path.dirname(os.path.abspath(self.current_file))

            # Open folder based on platform
            try:
                if platform.system() == "Windows":
                    os.startfile(folder_path)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", folder_path])
                else:  # Linux
                    subprocess.run(["xdg-open", folder_path])
            except Exception as e:
                messagebox.showerror("Error", f"Cannot open folder:\n{str(e)}")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = ExamCalendarApp()
    app.run()