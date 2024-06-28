import tkinter as tk
from tkinter import ttk, messagebox
import secrets
import string
import pandas as pd
import os


class PasswordGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Random Password Generator")
        self.root.geometry("400x400")
        self.root.resizable(False, False)
        self.root.configure(bg='#e6f7ff')  # Light blue background for the window

        # Create a style
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # Define custom styles with colors
        self.style.configure('TLabel', font=('Helvetica', 12), background='#e6f7ff',
                             foreground='#003366')  # Dark blue text
        self.style.configure('TButton', font=('Helvetica', 12), background='#80ccff',
                             foreground='#003366')  # Light blue button with dark blue text
        self.style.map('TButton', background=[('active', '#3399ff')])  # Darker blue on hover
        self.style.configure('TCheckbutton', font=('Helvetica', 12), background='#e6f7ff',
                             foreground='#003366')  # Dark blue text
        self.style.configure('TEntry', font=('Helvetica', 12))

        # Main frame with padding and background color
        main_frame = ttk.Frame(root, padding="10 10 10 10", style='TFrame')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Site Name Label and Entry
        self.site_label = ttk.Label(main_frame, text="Site Name:", style='TLabel')
        self.site_label.grid(row=0, column=0, pady=10, sticky=tk.W)

        self.site_var = tk.StringVar()
        self.site_entry = ttk.Entry(main_frame, textvariable=self.site_var, width=30, style='TEntry')
        self.site_entry.grid(row=0, column=1, pady=10, sticky=tk.E)

        # Length Label and Entry
        self.length_label = ttk.Label(main_frame, text="Password Length:", style='TLabel')
        self.length_label.grid(row=1, column=0, pady=10, sticky=tk.W)

        self.length_var = tk.StringVar()
        self.length_entry = ttk.Entry(main_frame, textvariable=self.length_var, width=10, style='TEntry')
        self.length_entry.grid(row=1, column=1, pady=10, sticky=tk.E)

        # Character Type Checkbuttons
        self.include_uppercase = tk.BooleanVar(value=True)
        self.include_digits = tk.BooleanVar(value=True)
        self.include_punctuation = tk.BooleanVar(value=True)

        self.uppercase_check = ttk.Checkbutton(main_frame, text="Include Uppercase", variable=self.include_uppercase,
                                               style='TCheckbutton')
        self.uppercase_check.grid(row=2, column=0, columnspan=2, sticky=tk.W)

        self.digits_check = ttk.Checkbutton(main_frame, text="Include Digits", variable=self.include_digits,
                                            style='TCheckbutton')
        self.digits_check.grid(row=3, column=0, columnspan=2, sticky=tk.W)

        self.punctuation_check = ttk.Checkbutton(main_frame, text="Include Punctuation",
                                                 variable=self.include_punctuation, style='TCheckbutton')
        self.punctuation_check.grid(row=4, column=0, columnspan=2, sticky=tk.W)

        # Generate Button
        self.generate_button = ttk.Button(main_frame, text="Generate Password", command=self.generate_password,
                                          style='TButton')
        self.generate_button.grid(row=5, column=0, columnspan=2, pady=20)

        # Save Button
        self.save_button = ttk.Button(main_frame, text="Save to Excel", command=self.save_to_excel, style='TButton')
        self.save_button.grid(row=6, column=0, columnspan=2, pady=10)

        # Result Entry (Read-Only)
        self.result_var = tk.StringVar()
        self.result_entry = ttk.Entry(main_frame, textvariable=self.result_var, state='readonly', width=30,
                                      style='TEntry')
        self.result_entry.grid(row=7, column=0, columnspan=2, pady=5)

    def generate_password(self):
        try:
            length = int(self.length_var.get())
            if length < 1:
                raise ValueError("Length must be at least 1")

            # Define the character set
            characters = string.ascii_lowercase
            if self.include_uppercase.get():
                characters += string.ascii_uppercase
            if self.include_digits.get():
                characters += string.digits
            if self.include_punctuation.get():
                characters += string.punctuation

            # Generate a secure random password
            password = ''.join(secrets.choice(characters) for _ in range(length))

            # Set the password in the result entry
            self.result_var.set(password)
        except ValueError as e:
            messagebox.showerror("Invalid Input", f"Error: {e}")

    def save_to_excel(self):
        site = self.site_var.get()
        password = self.result_var.get()

        if not site or not password:
            messagebox.showerror("Input Error", "Site name and password must not be empty.")
            return

        # Create a DataFrame
        df = pd.DataFrame([[site, password]], columns=['Site', 'Password'])

        # Check if the Excel file exists
        file_exists = os.path.isfile('passwords.xlsx')

        if file_exists:
            # Append the data to the existing file
            existing_df = pd.read_excel('passwords.xlsx')
            df = pd.concat([existing_df, df], ignore_index=True)

        # Save the DataFrame to an Excel file
        df.to_excel('passwords.xlsx', index=False)
        messagebox.showinfo("Success", "Password saved to passwords.xlsx")


if __name__ == "__main__":
    root = tk.Tk()
    app = PasswordGeneratorApp(root)
    root.mainloop()
