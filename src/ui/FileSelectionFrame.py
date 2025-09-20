from tkinter import filedialog
import customtkinter
import os

class FileSelectionFrame(customtkinter.CTkFrame):
    file_paths: dict[str, str|None] = {
        "transactions_file": None,
        "credits_file": None,
        "output_directory": None,
    }

    def __init__(self, master):
        super().__init__(master)
        self.grid_columnconfigure((0,1), weight=1)
        self.select_transaction_file_btn = customtkinter.CTkButton(self, text="Select Transaction File", command=self.select_transactions_file, width=200)
        self.select_transaction_file_btn.grid(row=1, column=0, padx=20, pady=20)
        self.transactions_file_label = customtkinter.CTkLabel(self, text="No file selected.", fg_color="transparent", corner_radius=6)
        self.transactions_file_label.grid(row=1, column=1, padx=20, pady=20)

        self.select_credits_file_btn = customtkinter.CTkButton(self, text="Select Credits File", command=self.select_credits_file, width=200)
        self.select_credits_file_btn.grid(row=2, column=0, padx=20, pady=20)
        self.credits_file_label = customtkinter.CTkLabel(self, text="No file selected.", fg_color="transparent", corner_radius=6)
        self.credits_file_label.grid(row=2, column=1, padx=20, pady=20)

        self.select_output_directory_btn = customtkinter.CTkButton(self, text="Select Output Directory", command=self.select_output_directory, width=200)
        self.select_output_directory_btn.grid(row=3, column=0, padx=20, pady=20)
        self.output_directory_label = customtkinter.CTkLabel(self, text="No directory selected.", fg_color="transparent", corner_radius=6)
        self.output_directory_label.grid(row=3, column=1, padx=20, pady=20)

    def select_transactions_file(self) -> None:
        filepath = filedialog.askopenfilename(
                title="Select a transaction file",
                filetypes=(("CSV files", "*.csv"), ("All files", "*.*"))
            )
        if filepath:
            self.file_paths["transactions_file"] = filepath
            filename = os.path.basename(filepath)
            self.transactions_file_label.configure(text=filename)

    def select_credits_file(self) -> None:
        filepath = filedialog.askopenfilename(
                title="Select a credits file",
                filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
            )
        if filepath:
            self.file_paths["credits_file"] = filepath
            filename = os.path.basename(filepath)
            self.credits_file_label.configure(text=filename)

    def select_output_directory(self) -> None:
        dirpath = filedialog.askdirectory(title="Select an output directory")
        if dirpath:
            self.file_paths["output_directory"] = dirpath
            self.output_directory_label.configure(text=dirpath)

    def get_file_paths(self) -> dict[str, str|None]:
      return self.file_paths
