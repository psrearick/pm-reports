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

class App(customtkinter.CTk):
    app_name: str = "Report Generator"

    def __init__(self):
        super().__init__()
        self.setup_ui()

    def setup_ui(self):
        self.title(self.app_name)
        self.geometry("800x500")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self.title_label = customtkinter.CTkLabel(self, text=self.app_name, fg_color="transparent", corner_radius=6)
        self.title_label.grid(row=0, column=0, padx=20, pady=(10, 0), sticky="ew", columnspan=2)

        self.file_selection_frame = FileSelectionFrame(self)
        self.file_selection_frame.grid(row=1, column=0, padx=10, pady=20, sticky="nesw")

        self.spacer = customtkinter.CTkLabel(self, text="", height=20)
        self.spacer.grid(row=2, column=0)

        self.progress_bar = customtkinter.CTkProgressBar(master=self, mode="determinate")
        self.progress_bar.grid(row=3, column=0, pady=(80,20), padx=10, sticky="ew")
        self.progress_bar.set(0)

        self.generate_btn = customtkinter.CTkButton(self, text="Generate Report", command=self.generate_reports, hover_color="#007a55", fg_color="#009966")
        self.generate_btn.grid(row=4, column=0, padx=20, pady=20)

    def generate_reports(self):
        paths = self.file_selection_frame.get_file_paths()
        print(paths)
        self.progress_bar.set(0)
        self.update_progress()

    def update_progress(self):
        current_value = self.progress_bar.get()
        if current_value < 1:
            self.progress_bar.set(current_value + 0.1)
            self.after(500, self.update_progress)
        else:
            print("Progress complete!")
