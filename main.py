import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox, simpledialog
import random
import os
import json
from datetime import datetime
import docx
from docx.shared import Pt, Inches
from question_paper_generator_ import QuestionPaperGenerator, save_question_paper
from constants import options


class QuestionPaperApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Question Paper Generator")
        self.root.geometry("800x600")  # sets the window size to 800x600 pixels
        self.root.resizable(True, True)  # makes the window resizable in both width and height
        self.questions = []
        self.login_data_file = "login_data.json"
        self.current_user = None

        script_dir = os.path.dirname(__file__)
        logo_path = os.path.join(script_dir, "logo.png")
        logo2_path = os.path.join(script_dir, "logo2.png")

        self.logo_image = tk.PhotoImage(file=logo_path)
        self.logo2_image = tk.PhotoImage(file=logo2_path)

        self.create_login_page()

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def create_login_page(self):
        self.clear_window()
        login_frame = tk.Frame(self.root, bg="#f0f0f0")
        login_frame.place(relx=0.5, rely=0.5, anchor="center")

        logo_label = tk.Label(login_frame, image=self.logo_image)
        logo_label.image = self.logo_image  # keep a reference to prevent garbage collection
        logo_label.pack(pady=10)

        tk.Label(login_frame, text="Login", font=("Arial", 24, "bold"), fg="#00698f").pack(pady=10)
        tk.Label(login_frame, text="Username", font=("Arial", 16)).pack()
        self.username_entry = tk.Entry(login_frame, font=("Arial", 16), width=30)
        self.username_entry.pack()
        tk.Label(login_frame, text="Password", font=("Arial", 16)).pack()
        self.password_entry = tk.Entry(login_frame, font=("Arial", 16), show='*', width=30)
        self.password_entry.pack()

        login_button = ttk.Button(login_frame, text="Login", command=self.login, style="Login.TButton")
        login_button.pack(pady=5)

        register_button = ttk.Button(login_frame, text="Register", command=self.create_register_page,
                                     style="Login.TButton")
        register_button.pack(pady=5)

        self.root.style = ttk.Style()
        self.root.style.configure("Login.TButton", font=("Arial", 16), foreground="#00698f", background="#ccefff",
                                  relief="raised", borderwidth=2)

    def create_register_page(self):

        self.clear_window()
        register_frame = tk.Frame(self.root, bg="#f0f0f0")
        register_frame.place(relx=0.5, rely=0.5, anchor="center")

        logo_label = tk.Label(register_frame, image=self.logo_image)
        logo_label.image = self.logo_image  # keep a reference to prevent garbage collection
        logo_label.pack(pady=10)

        tk.Label(register_frame, text="Register", font=("Arial", 24, "bold"), fg="#00698f").pack(pady=10)
        tk.Label(register_frame, text="Name:", font=("Arial", 16)).pack()
        self.new_name_entry = tk.Entry(register_frame, font=("Arial", 16), width=30)
        self.new_name_entry.pack()
        tk.Label(register_frame, text="ID:", font=("Arial", 16)).pack()
        self.new_id_entry = tk.Entry(register_frame, font=("Arial", 16), width=30)
        self.new_id_entry.pack()
        tk.Label(register_frame, text="Username:", font=("Arial", 16)).pack()
        self.new_username_entry = tk.Entry(register_frame, font=("Arial", 16), width=30)
        self.new_username_entry.pack()
        tk.Label(register_frame, text=" Password:", font=("Arial", 16)).pack()
        self.new_password_entry = tk.Entry(register_frame, font=("Arial", 16), show='*', width=30)
        self.new_password_entry.pack()

        register_button = ttk.Button(register_frame, text="Register", command=self.register, style="Register.TButton")
        register_button.pack(pady=5)

        back_button = ttk.Button(register_frame, text="Back to Login", command=self.create_login_page,
                                 style="Register.TButton")
        back_button.pack(pady=5)

        self.root.style = ttk.Style()
        self.root.style.configure("Register.TButton", font=("Arial", 16), foreground="#00698f", background="#ccefff",
                                  relief="raised", borderwidth=2)

    def login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()
        with open(self.login_data_file, 'r') as f:
            users = json.load(f)
        if username in users and users[username]["password"] == password:
            self.current_user = username
            self.log_login_activity(username)
            self.create_main_page()
        else:
            messagebox.showerror("Error", "Invalid credentials")

    def log_login_activity(self, username):
        with open("login_activity.txt", "a") as f:
            f.write(f"{username} logged in at {datetime.now()}\n")

    def register(self):
        name = self.new_name_entry.get()
        id = self.new_id_entry.get()
        username = self.new_username_entry.get()
        password = self.new_password_entry.get()
        if name and id and username and password:
            with open(self.login_data_file, 'r') as f:
                users = json.load(f)
            if username not in users:
                users[username] = {"password": password, "name": name, "id": id}
                with open(self.login_data_file, 'w') as f:
                    json.dump(users, f)
                messagebox.showinfo("Success", "Registration successful")
                self.create_login_page()
            else:
                messagebox.showerror("Error", "Username already exists")
        else:
            messagebox.showerror("Error", "Please fill in all fields")

    def create_main_page(self):
        self.clear_window()
        main_frame = tk.Frame(self.root, bg="#f0f0f0")
        main_frame.place(relx=0.5, rely=0.5, anchor="center")

        logo_label = tk.Label(main_frame, image=self.logo_image)
        logo_label.image = self.logo_image  # keep a reference to prevent garbage collection
        logo_label.pack(pady=10)

        tk.Label(main_frame, text="Question Paper Generator", font=("Arial", 24, "bold"), fg="#00698f").pack(pady=10)

        select_file_button = ttk.Button(main_frame, text="Select Question Bank", command=self.select_file,
                                        style="Main.TButton")
        select_file_button.pack(pady=5)

        generate_button = ttk.Button(main_frame, text="Generate Questions", command=self.generate_questions_page,
                                     style="Main.TButton")
        generate_button.pack(pady=5)

        logout_button = ttk.Button(main_frame, text="Logout", command=self.logout, style="Main.TButton")
        logout_button.pack(pady=5)

        self.root.style = ttk.Style()
        self.root.style.configure("Main.TButton", font=("Arial", 16), foreground="#00698f", background="#ccefff",
                                  relief="raised", borderwidth=2)

    def select_file(self):
        # Open a file dialog to select a .docx file
        self.file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])

        if self.file_path:
            # Create an instance of QuestionPaperGenerator
            generator = QuestionPaperGenerator(self.file_path, None)
            # Use the correct method to read questions
            self.questions = generator.read_questions_from_word(self.file_path)

            messagebox.showinfo("Success", "File loaded successfully")

    def show_unit_frame(self):
        if self.option_var.get() == "CIE":
            self.unit_frame.pack()
        else:
            self.unit_frame.pack_forget()

    def generate_questions_page(self):
        if not self.questions:
            messagebox.showerror("Error", "No questions loaded")
            return
        self.clear_window()
        generate_frame = tk.Frame(self.root, bg="#f0f0f0")
        generate_frame.place(relx=0.5, rely=0.5, anchor="center")

        logo_label = tk.Label(generate_frame, image=self.logo2_image)
        logo_label.image = self.logo2_image  # keep a reference to prevent garbage collection
        logo_label.pack(pady=10)

        tk.Label(generate_frame, text="Select Option", font=("Arial", 24, "bold"), fg="#00698f").pack(pady=10)

        self.option_var = tk.StringVar()
        self.option_var.set("SEE")  # default value

        button_frame = tk.Frame(generate_frame)
        button_frame.pack(pady=10)

        see_button = ttk.Button(button_frame, text="SEE", command=lambda: self.option_var.set("SEE"),
                                style="Option.TButton")
        see_button.pack(side=tk.LEFT, padx=10)

        cie_button = ttk.Button(button_frame, text="CIE", style="Option.TButton")
        cie_button.pack(side=tk.LEFT, padx=10)


        #tick box of units
        self.unit_frame = tk.Frame(generate_frame, bg="#E8E6F5", highlightbackground="#00698f", highlightthickness=2)
        self.unit_frame.pack_forget()

        self.unit_vars = [tk.IntVar() for _ in range(5)]

        unit_checkbox_frame = tk.Frame(self.unit_frame, bg="#E8E6F5")
        unit_checkbox_frame.pack(pady=20, padx=20)

        for i, var in enumerate(self.unit_vars):
            checkbox = tk.Checkbutton(unit_checkbox_frame, text=f"Unit {i + 1}", variable=var, font=("Arial", 14),
                                      bg="#E8E6F5", selectcolor="#ccefff")
            checkbox.pack(side=tk.LEFT, padx=10)

        cie_button.config(command=lambda: [self.option_var.set("CIE"), self.show_unit_frame()])
        see_button.config(command=lambda: [self.option_var.set("SEE"), self.show_unit_frame()])





        tk.Label(generate_frame, text="Course Name", font=("Arial", 16)).pack()
        self.course_name_entry = tk.Entry(generate_frame, font=("Arial", 16), width=30)
        self.course_name_entry.pack()

        tk.Label(generate_frame, text="Course Code", font=("Arial", 16)).pack()
        self.course_code_entry = tk.Entry(generate_frame, font=("Arial", 16), width=30)
        self.course_code_entry.pack()

        tk.Label(generate_frame, text="Examination Date", font=("Arial", 16)).pack()
        self.examination_date_entry = tk.Entry(generate_frame, font=("Arial", 16), width=30)
        self.examination_date_entry.pack()

        self.Custom_frame = tk.Frame(generate_frame)
        self.Custom_frame.pack_forget()

        tk.Label(self.Custom_frame, text="SAQs", font=("Arial", 16)).pack(side=tk.LEFT)
        self.saq_num_entry = tk.Entry(self.Custom_frame, font=("Arial", 16), width=10)
        self.saq_num_entry.pack(side=tk.LEFT)

        tk.Label(self.Custom_frame, text="LAQs", font=("Arial", 16)).pack(side=tk.LEFT)
        self.laq_num_entry = tk.Entry(self.Custom_frame, font=("Arial", 16), width=10)
        self.laq_num_entry.pack(side=tk.LEFT)

        self.root.style = ttk.Style()
        self.root.style.configure("Option.TButton", font=("Arial", 16), foreground="#00698f", background="#ccefff",
                                  relief="raised", borderwidth=2)
        self.root.style.configure("Generate.TButton", font=("Arial", 16), foreground="#00698f", background="#ccefff",
                                  relief="raised", borderwidth=2)

        generate_button = ttk.Button(generate_frame, text="Generate", command=self.generate_questions,
                                     style="Generate.TButton")
        generate_button.pack(pady=5)
        BACK_button = ttk.Button(generate_frame, text="Back", command=self.create_main_page,
                                 style="Generate.TButton")
        BACK_button.pack(pady=5)

        def show_unit_frame(self):
            if self.option_var.get() == "CIE":
                self.unit_frame.pack()
            else:
                self.unit_frame.pack_forget()

        cie_button.config(command=lambda: [self.option_var.set("CIE"), self.show_unit_frame()])
        see_button.config(command=lambda: [self.option_var.set("SEE"), self.show_unit_frame()])

    def generate_questions(self):
        option = self.option_var.get()
        course_name = self.course_name_entry.get()
        course_code = self.course_code_entry.get()
        examination_date = self.examination_date_entry.get()

        # Check if all fields are filled
        if not option or not course_name or not course_code or not examination_date:
            messagebox.showerror("Error", "Please fill in all fields")
            return

        # Validate selected option
        if option not in [opt["name"] for opt in options]:
            messagebox.showerror("Error", "Invalid option selected")
            return

        # Handle custom option
        if option == "Custom":
            try:
                saq_num = int(self.saq_num_entry.get())
                laq_num = int(self.laq_num_entry.get())
            except ValueError:
                messagebox.showerror("Error", "Invalid SAQs or LAQs number")
                return
        else:
            for opt in options:
                if opt["name"] == option:
                    saq_num = opt["saq_num"]
                    laq_num = opt["laq_num"]
                    break

        # Ensure 3 units are selected for CIE
        selected_units = [i + 1 for i, var in enumerate(self.unit_vars) if var.get() and option == "CIE"]
        if option == "CIE" and len(selected_units) != 3:
            messagebox.showerror("Error", "Please select exactly 3 units")
            return

        # Initialize the QuestionPaperGenerator with required parameters
        generator = QuestionPaperGenerator(self.file_path, option, saq_num, laq_num)

        # Pass exam type ('SEE' or 'CIE') when calling generate_question_paper
        if option == "CIE":
            saqs, laqs = generator.generate_question_paper(selected_units, exam_type="CIE")
        else:
            saqs, laqs = generator.generate_question_paper(exam_type="SEE")

        # Prompt the user to select a save location
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if file_path:
            try:
                # Save the generated question paper to the specified file
                save_question_paper(saqs, laqs, file_path, course_name, course_code, examination_date, option)
                messagebox.showinfo("Success", "Questions saved to " + os.path.basename(file_path))
                self.create_main_page()  # Reset the UI or go back to main page
            except Exception as e:
                messagebox.showerror("Error", str(e))
        else:
            messagebox.showerror("Error", "No file selected")

    def logout(self):
        self.log_logout_activity(self.current_user)
        self.current_user = None
        self.create_login_page()

    def log_logout_activity(self, username):
        with open("login_activity.txt", "a") as f:
            f.write(f"{username} logged out at {datetime.now()}\n")

    def clear_window(self):
        for widget in self.root.winfo_children():
            widget.destroy()


if __name__ == "__main__":
    print("Script is running...")
    if not os.path.exists("login_data.json"):
        with open("login_data.json", 'w') as f:
            json.dump({},f)

    root = tk.Tk()
    app = QuestionPaperApp(root)
    root.mainloop()