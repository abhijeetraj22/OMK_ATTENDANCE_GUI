import tkinter as tk
from tkinter import messagebox, ttk

class TextReplaceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Text Replace Tool")

        # --- Center window on screen ---
        window_width = 600
        window_height = 400
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # ✅ Use grid layout to manage space
        self.grid_rowconfigure(0, weight=1)  # Text area expands
        self.grid_columnconfigure(0, weight=1)

        # --- Frame for text + scrollbars ---
        text_frame = tk.Frame(self)
        text_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 0))

        # Vertical + Horizontal scrollbars
        v_scroll = tk.Scrollbar(text_frame, orient="vertical")
        v_scroll.pack(side="right", fill="y")

        h_scroll = tk.Scrollbar(text_frame, orient="horizontal")
        h_scroll.pack(side="bottom", fill="x")

        # Text widget
        self.text_box = tk.Text(
            text_frame, wrap="none", font=("Arial", 12),
            xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set
        )
        self.text_box.pack(side="left", expand=True, fill="both")

        v_scroll.config(command=self.text_box.yview)
        h_scroll.config(command=self.text_box.xview)

        # --- Button frame (fixed at bottom, centered) ---
        button_frame = tk.Frame(self)
        button_frame.grid(row=1, column=0, pady=10)

        replace_btn = tk.Button(
            button_frame, text="Replace Text", command=self.replace_text,
            font=("Arial", 12, "bold"), bg="lightblue",width=12
        )
        replace_btn.pack(side="left", padx=20)

        exit_btn = tk.Button(
            button_frame, text="Exit", command=self.show_final_frame,
            font=("Arial", 12, "bold"), bg="red", fg="white",width=12
        )
        exit_btn.pack(side="left", padx=20)

        # Handle "X" button click also
        self.protocol("WM_DELETE_WINDOW", self.show_final_frame)

    def decode_escape_sequences(self, s: str) -> str:
        """Convert \n, \t, \\ into real characters"""
        return s.encode('utf-8').decode('unicode_escape')

    def replace_text(self):
        text = self.text_box.get("1.0", tk.END).rstrip()
        if not text:
            messagebox.showwarning("Warning", "Please paste some text first!")
            return

        # Popup dialog
        dialog = tk.Toplevel(self)
        dialog.title("Find & Replace")

        # --- Center popup ---
        window_width = 350
        window_height = 220
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))
        dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")
        dialog.resizable(False, False)
        dialog.grab_set()

        # Dropdown
        tk.Label(dialog, text="Quick Replace (Dropdown):").pack(pady=2)
        options = [
            "None",
            "\\n → , ",
            ",  → \\n",
            "\\t → space",
            "  → _"
        ]
        combo = ttk.Combobox(dialog, values=options, state="readonly")
        combo.current(0)
        combo.pack(pady=5)

        # Manual entries
        tk.Label(dialog, text="Custom Find:").pack(pady=2)
        find_entry = tk.Entry(dialog, width=30)
        find_entry.pack(pady=2)

        tk.Label(dialog, text="Custom Replace With:").pack(pady=2)
        replace_entry = tk.Entry(dialog, width=30)
        replace_entry.pack(pady=2)

        def do_replace():
            new_text = text

            # Apply dropdown replacement
            dropdown_value = combo.get()
            if dropdown_value != "None":
                find_str, replace_str = dropdown_value.split(" → ")
                find_str = self.decode_escape_sequences(find_str)
                replace_str = self.decode_escape_sequences(replace_str)
                new_text = new_text.replace(find_str, replace_str)

            # Apply manual replacement
            find_str = find_entry.get()
            replace_str = replace_entry.get()
            if find_str:
                find_str = self.decode_escape_sequences(find_str)
                replace_str = self.decode_escape_sequences(replace_str)
                new_text = new_text.replace(find_str, replace_str)

            # Update text box
            self.text_box.delete("1.0", tk.END)
            self.text_box.insert(tk.END, new_text)
            dialog.destroy()

        tk.Button(dialog, text="OK", command=do_replace).pack(pady=10)

    # ✅ New function for final thank-you screen
    def show_final_frame(self):
        for widget in self.winfo_children():
            widget.destroy()

        self.final_frame = tk.Frame(self)
        self.final_frame.pack(fill='both', expand=True)

        tk.Label(
            self.final_frame,
            text="🎉 Thank you for using this Text Replace Tool!",
            font=("Arial", 12, "bold"), fg='green'
        ).pack(pady=60)

        tk.Button(
            self.final_frame, text="Exit", command=self.destroy,
            font=("Arial", 12, "bold"), width=20, bd=0, bg="red", fg="white"
        ).pack(pady=10)


if __name__ == "__main__":
    app = TextReplaceApp()
    app.mainloop()
