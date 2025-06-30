import tkinter as tk
from mi_app.gui import GoogleToDocApp

def main():
    """Main entry point for the application"""
    root = tk.Tk()
    app = GoogleToDocApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()