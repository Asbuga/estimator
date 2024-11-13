""" 
Invokes main window app.
"""
from generate_word_from_excel.libs import Window


def main():
    app = Window()
    app.mainloop()


if __name__ == "__main__":
    main()
