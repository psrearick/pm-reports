import customtkinter

class App(customtkinter.CTk):
    app_name: str = "Report Generator"

    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.file_paths = {}

    def setup_ui(self):
        self.title(self.app_name)
        self.geometry("400x200")
        self.grid_columnconfigure((0,1), weight=1)
        self.title = customtkinter.CTkLabel(self, text=self.app_name, fg_color="transparent", corner_radius=6)
        self.title.grid(row=0, column=0, padx=20, pady=(10, 0), sticky="ew", columnspan=2)

        # file selection section
        # output directory selection

        self.generate_btn = customtkinter.CTkButton(self, text="Generate Report", command=self.generate_reports, hover_color="#007a55", fg_color="#009966")
        self.generate_btn.grid(row=1, column=0, padx=20, pady=20)

        self.close_btn = customtkinter.CTkButton(self, text="Exit", command=self.close_window)
        self.close_btn.grid(row=1, column=1, padx=20, pady=20)

        # progress display

    def select_transactions_file(self):
        pass

    def select_credits_file(self):
        pass

    def select_config_file(self):
        pass

    def validate_inputs(self) -> bool:
        pass

    def generate_reports(self):
        pass

    def show_progress(self, message: str, percent: int):
        pass

    def close_window(self):
        self.destroy()
