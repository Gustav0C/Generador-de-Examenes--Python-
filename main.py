from gui import GeneradorExamenes

import os
import tkinter as tk

if __name__ == "__main__":
    root = tk.Tk()
    app = GeneradorExamenes.SoftwareExamenAdmision(root)
    root.mainloop()
