import tkinter as tk
from tkinter import messagebox
import subprocess
import sys


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ARM SOFT")

        # --- Center Window ---
        window_width = 600
        window_height = 400
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.resizable(False, False)

        # --- Fonts ---
        self.button_font = ("Arial", 14, "bold")

        # --- Buttons ---
        tk.Label(self, text="Choose a Tool", font=("Arial", 16, "bold"), fg="blue").pack(pady=20)

        tk.Button(
            self, text="Excel Task Tool",
            command=self.run_excel_app,
            font=self.button_font, bg="#258EFF", fg="white",
            width=20, height=2
        ).pack(pady=15)

        tk.Button(
            self, text="Text Replace Tool",
            command=self.run_textreplace_app,
            font=self.button_font, bg="green", fg="white",
            width=20, height=2
        ).pack(pady=15)

        # Exit
        tk.Button(
            self, text="Exit",
            command=self.show_final_frame,
            font=self.button_font, bg="red", fg="white",
            width=20, height=2
        ).pack(pady=15)

        # Handle "X" button click also
        self.protocol("WM_DELETE_WINDOW", self.show_final_frame)

    def run_and_return(self, program_name):
        try:
            self.withdraw()

            if program_name.endswith(".exe"):
                # Run exe directly
                process = subprocess.Popen([program_name])
            else:
                # Run python script with current interpreter
                process = subprocess.Popen([sys.executable, program_name])

            process.wait()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to start {program_name}:\n{e}")
        finally:
            self.deiconify()


    def run_excel_app(self):
        self.run_and_return("OMK_ATTENDANCE_GUI.exe")

    def run_textreplace_app(self):
        self.run_and_return("replaceText.exe")

    def show_final_frame(self):
        for widget in self.winfo_children():
            widget.destroy()

        self.final_frame = tk.Frame(self)
        self.final_frame.pack(fill='both', expand=True)

        tk.Label(
            self.final_frame,
            text="🎉 Thank you for using this ARM Software!",
            font=("Arial", 12, "bold"), fg='green'
        ).pack(pady=60)

        tk.Button(
            self.final_frame, text="Exit", command=self.destroy,
            font=("Arial", 12, "bold"), width=20, bd=0, bg="red", fg="white"
        ).pack(pady=10)


if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
