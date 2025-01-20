""" 
Invokes main window app.
"""
from .libs import Window


def main():
    app = Window()
    app.mainloop()


if __name__ == "__main__":
    main()
