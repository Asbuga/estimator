import os
import webbrowser
import tkinter as tk
from tkinter import filedialog

from create_word import create_word_file


class BaseButton(tk.Button):
    def __init__(self, master, text, command=None, bg=None, **kwargs):
        super().__init__(master, text=text, command=command, bg=bg, **kwargs)
        self.configure(bg="#ffcccb", fg="black", font=("Arial", 11), height=2, width=20)
        self.pack(pady=1, padx=5, fill=tk.X)

    def change_color(self, result=None):
        if result:
            self.config(bg="lightgreen")


class Window(tk.Tk):
    def __init__(self):
        super().__init__()

        # Window settings.
        self.geometry("320x285")
        self.title("Generate Word from Excel")
        self.resizable(False, False)
        self.configure(background="grey")
        self.attributes("-topmost", True)

        # Buttons settings.
        self.select_excel_button = BaseButton(
            self, text="Вибрати файл Excel", command=self.select_excel
        )

        self.select_word_button = BaseButton(
            self, text="Вибрати файл Word", command=self.select_word
        )

        self.select_path_button = BaseButton(
            self, text="Зберегти результат в ...", command=self.select_path
        )

        self.generate_button = BaseButton(
            self, text="Згенерувати файли", command=self.generate_files
        )

        self.about = BaseButton(
            self,
            text="Про програму",
            command=lambda: webbrowser.open("github.com"),
        )
        self.about.configure(bg="lightgreen")

        self.exit_app = BaseButton(
            self,
            text="Вихід",
            command=lambda: self.destroy(),
        )
        self.exit_app.configure(bg="lightgreen")

        # Variable parameters.
        self.initialdir = "~/Desktop/"
        self.excel_file = None
        self.word_file = None
        self.output_path = None

    def select_excel(self):
        self.excel_file = filedialog.askopenfilename(initialdir=self.initialdir, title="Виберіть Excel файл з таблицею для вставки:", filetypes=(("EXCEL files", "*.xlsx"), ("EXCEL files", "*.xls"),))
        self.select_excel_button.change_color(self.excel_file)
        self.generate_color()

    def select_word(self):
        self.word_file = filedialog.askopenfilename(initialdir=self.initialdir, title="Виберіть Word файл шаблону:", filetypes=(("DOCX files", "*.docx"),))
        self.select_word_button.change_color(self.word_file)
        self.generate_color()

    def select_path(self):
        self.output_path = filedialog.askdirectory(initialdir=self.initialdir, title="Виберіть папку для збереження згенерованих документів:")
        self.select_path_button.change_color(self.output_path)
        self.generate_color()

    def generate_color(self):
        if all([self.excel_file, self.word_file, self.output_path]):
            self.generate_button.config(bg="lightgreen")

    def generate_files(self):
        if self.output_path:
            create_word_file(
                table=self.excel_file, template=self.word_file, path_=self.output_path
            )
            os.startfile(self.output_path)


if __name__ == "__main__":
    app = Window()
    app.mainloop()
