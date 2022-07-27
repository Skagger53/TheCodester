import re
import tkinter as tk
from tkinter import messagebox
from PIL import ImageTk, Image
import pyperclip

class MainWindow:
    def __init__(self):
        self.codes = [] # Used for code output (list)
        self.new_list = []  # Used to correct list of codes (no dupes, etc.)
        self.window_width, self.window_height = 660, 830 # Sets program size

        self.main_window = tk.Tk()
        self.main_window.resizable(False, False)

        # Sets window location in the middle of the screen (offset up by 5% of window height -- cosmetic choice)
        screen_width = self.main_window.winfo_screenwidth()
        screen_height = self.main_window.winfo_screenheight()

        x = (screen_width/2) - (self.window_width/2)
        y = (screen_height/2) - (self.window_height/2) - screen_height/20

        self.main_window.geometry('%dx%d+%d+%d' % (self.window_width, self.window_height, x, y))

        # self.main_window.geometry("660x830")
        self.main_window.title("The Codester is here!")

        # Top frame contains label for text widget (user text input) and extract_button
        self.top_frame = tk.Frame(self.main_window, pady = 10)
        self.top_frame.grid(row = 0, column = 1, rowspan = 4, columnspan = 2, padx = 10, pady = (15, 0))

        # Frame for the Codester himself!
        self.test_frame = tk.Frame(self.main_window, padx = 10)
        self.test_frame.grid(row = 0, rowspan = 4, column = 0, sticky = "e")

        # Opening and resizing the Codester
        the_codester = Image.open("TheCodesterHimself.png")
        the_codester = ImageTk.PhotoImage(the_codester.resize((187, 205), Image.Resampling.LANCZOS))
        label = tk.Label(self.test_frame, image = the_codester)
        label.pack()

        # Label for user input text widget
        self.get_text_lbl = tk.Label(self.top_frame, text="Enter text here. Kick back, relax, and leave it to the Codester!")
        self.get_text_lbl.grid(column=1, row=0, sticky="W", pady = (0, 10))

        # Setting up text widget for user input
        self.input_text = tk.Text(self.top_frame, height=25, width=48)
        self.input_text.grid(row=1, column=1, padx=(0, 10))

        # Sets up scrollbar for text widget for user input
        self.input_text_scrollbar = tk.Scrollbar(self.top_frame, command=self.input_text.yview, orient="vertical")
        self.input_text_scrollbar.grid(row=1, column=2, sticky="nse")
        self.input_text.configure(yscrollcommand=self.input_text_scrollbar.set)

        # Extract button
        self.extract_button = tk.Button(self.top_frame, text = "Evaluate", command = self.evaluate_data)
        self.extract_button.grid(column = 1, row = 2, pady = (15, 0))

        # Bottom frame contains label for output, output text widget, and the copy to clipboard button.
        self.bottom_frame = tk.Frame(self.main_window, padx = 10, pady = 10)
        self.bottom_frame.grid(row=4, column=0, rowspan=2, columnspan=3, padx = 10, pady = (0, 15))

        # Button to copy output codes to clipboard. Disabled by default and only enabled when codes are entered.
        self.copy_button = tk.Button(self.bottom_frame, text = "Copy codes to clipboard", command = self.copy_to_clipboard)
        self.copy_button.grid(column=1, row=6, pady=(15, 0))
        self.copy_button["state"] = "disabled"

        # Label over output text widget. This starts out as empty and only contains data if (1) codes were found or (2) text was searched and zero codes were found
        self.codes_output_label = tk.Label(self.bottom_frame, text="")
        self.codes_output_label.grid(column=0, columnspan=2, row=4, sticky="W", pady = (0, 10))

        # Text widget for codes output. This is disabled so that the user cannot type in it. (It must be enabled and disabled each time the program edits its contents)
        self.codes_output = tk.Text(self.bottom_frame, height = 9, width = 72, state = "disabled")
        self.codes_output.grid(column=1, row=5, pady = (0, 15))

        # Text widget for code output scroll bar
        self.codes_output_scrollbar = tk.Scrollbar(self.bottom_frame, command=self.codes_output.yview, orient="vertical")
        self.codes_output_scrollbar.grid(column=2, row=5, sticky="nse")
        self.input_text.configure(yscrollcommand=self.input_text_scrollbar.set)

        # Exit button
        self.exit_button = tk.Button(self.main_window, text = "Exit", command = self.confirm_exit)
        self.exit_button.grid(column = 1, row = 7, columnspan = 2, sticky = "e")

        self.main_window.mainloop()

    # Confirms that the user wants to exit
    def confirm_exit(self):
        self.button_press(self.exit_button) # Animates exit button
        if messagebox.askokcancel("You're not leaving, are you? ☹", "Are you sure you want to quit?\n\nThe Codester will miss you...") == True: self.main_window.destroy()

    # Manually depresses and then elevates button. This is cosmetic only; otherwise the buttons do not visually change when clicked.
    def button_press(self, button_to_raise):
        button_to_raise.config(relief="sunken")
        self.main_window.after(220, lambda: button_to_raise.config(relief="raised"))

    # Copies codes to the clipboard. Codes are already in text (not list) format.
    def copy_to_clipboard(self):
        pyperclip.copy(self.codes)
        self.button_press(self.copy_button) # Animates copy button

    # Takes the self.list list and ensures that it does not have any duplicates. Also sorts and converts to string with comma delimiters.
    def no_dups(self):
        # Checks for duplicates. From current point in list iteration, if a code is unique, it appears one time. If more than once, it's skipped (and will be caught on a later iteration).
        for i, code in enumerate(self.codes):
            if self.codes[i:].count(code) == 1: self.new_list.append(code)

        self.new_list.sort()
        self.new_list = ", ".join(self.new_list)

        return self.new_list

    # Called when user clicks "Evaluate" button
    def evaluate_data(self):
        user_input = self.input_text.get("1.0", "end-1c") # Obtains user's input (minus one character to avoid final hard return)

        # When the user enters nothing, codes_output text widget and label are reset to empty (e.g., if it had codes in it from a previous evaluation) and copy button is disabled
        if user_input == "":
            self.codes_output.config(state = "normal")
            self.codes_output.delete("1.0", "end")
            self.codes_output_label.config(text="")
            self.codes_output.config(state="disabled")

            self.copy_button["state"] = "disabled"
            return

        # Input is not empty, so it's evaluated here with a regular expression
        self.codes = re.findall("[A-Z]\d+.?\d+", user_input)

        if len(self.codes) != 0:
            self.codes = self.no_dups() # Removes duplicates, sorts, converts to string with comma delimiters

            # Enters codes found into the text widget. Must enable so the program can use it at all and disable so the user can't change the text.
            self.codes_output.config(state="normal")
            self.codes_output.delete("1.0", "end")
            self.codes_output.insert("insert", self.codes)
            self.codes_output.config(state="disabled")

            self.copy_button["state"] = "normal" # User now needs to be able to copy the output

            self.codes_output_label.config(text = "Who's a good boy? Who's a good boy? The Codester's a good boy!")

        else: # If no codes were found
            self.copy_button["state"] = "disable"

            # Enters text to say that no codes were found (but copy button is disabled)
            self.codes_output.config(state="normal")
            self.codes_output.delete("1.0", "end")
            self.codes_output.insert("insert", "(No ICD-10 codes found)")
            self.codes_output.config(state="disabled")

            self.codes_output_label.config(text="Sad Codester is sad...")

        self.button_press(self.extract_button) # Animates Extract button

window = MainWindow()