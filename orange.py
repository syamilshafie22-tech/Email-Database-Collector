
import tkinter as tk
from tkinter import messagebox


# Submit function
def submit_form():
    title = title_entry.get()
    email = email_entry.get()
    passkey = passkey_entry.get()

    date = date_entry.get()
    time = time_entry.get()
    price = price_entry.get()
    date2 = date_entry2.get()
    time2 = time_entry2.get()
    price2 = price_entry2.get()

    if not all([title, email, passkey, date, time, price]):
        messagebox.showwarning(
            "Warning",
            "Please fill in all fields."
        )
        return

    messagebox.showinfo(
        "Submission Successful",
        f"Title: {title}\n"
        f"Email: {email}\n"
        f"Passkey: ********\n\n"
        f"\nInfo 1:\n"
        f"Date: {date}\n"
        f"Time: {time}\n"
        f"Price: {price}"
        f"\nInfo 2:\n"
        f"Date: {date2}\n"
        f"Time: {time2}\n"
        f"Price: {price2}"
    )
def copy_info():
    if copy_var.get():
        date_entry2.delete(0, tk.END)
        date_entry2.insert(0, date_entry.get())

        time_entry2.delete(0, tk.END)
        time_entry2.insert(0, time_entry.get())

        price_entry2.delete(0, tk.END)
        price_entry2.insert(0, price_entry.get())

# Main window
root = tk.Tk()
root.title("Orange Application")
root.geometry("650x550")
root.resizable(False, False)


# Application heading
heading = tk.Label(
    root,
    text="Orange Application",
    font=("Arial", 18, "bold")
)
heading.pack(pady=20)


# USER INFORMATION SECTION
user_frame = tk.LabelFrame(
    root,
    text="User Information",
    font=("Arial", 11, "bold"),
    padx=10,
    pady=10
)
user_frame.pack(padx=20, fill="x")


# Three columns
for i in range(3):
    user_frame.columnconfigure(i, weight=1)


# Title
tk.Label(user_frame, text="Title").grid(
    row=0, column=0, pady=5
)
title_entry = tk.Entry(user_frame, width=18)
title_entry.grid(row=1, column=0, padx=5)


# Email
tk.Label(user_frame, text="Email").grid(
    row=0, column=1, pady=5
)
email_entry = tk.Entry(user_frame, width=18)
email_entry.grid(row=1, column=1, padx=5)


# Passkey
tk.Label(user_frame, text="Passkey").grid(
    row=0, column=2, pady=5
)
passkey_entry = tk.Entry(
    user_frame,
    width=18,
    show="*"
)
passkey_entry.grid(row=1, column=2, padx=5)

# INFO SECTION
info_frame = tk.LabelFrame(
    root,
    text="Info 1",
    font=("Arial", 11, "bold"),
    padx=10,
    pady=10
)
info_frame.pack(padx=20, pady=15, fill="x")


# Date
tk.Label(info_frame, text="Date").grid(
    row=0, column=0, padx=5, pady=5
)
date_entry = tk.Entry(info_frame, width=10)
date_entry.grid(row=0, column=1, padx=5)


# Time
tk.Label(info_frame, text="Time").grid(
    row=0, column=2, padx=5, pady=5
)
time_entry = tk.Entry(info_frame, width=10)
time_entry.grid(row=0, column=3, padx=5)


# Price
tk.Label(info_frame, text="Price").grid(
    row=0, column=4, padx=5, pady=5
)
price_entry = tk.Entry(info_frame, width=10)
price_entry.grid(row=0, column=5, padx=5)

# INFO 2 SECTION
info_frame2 = tk.LabelFrame(
    root,
    text="Info 2",
    font=("Arial", 11, "bold"),
    padx=10,
    pady=10
)
info_frame2.pack(padx=20, pady=15, fill="x")


# Checkbox
copy_var = tk.BooleanVar()

copy_checkbox = tk.Checkbutton(
    info_frame2,
    variable=copy_var,
    command=copy_info
)
copy_checkbox.grid(row=0, column=0, padx=5)


# Date
tk.Label(info_frame2, text="Date").grid(
    row=0, column=1, padx=5, pady=5
)
date_entry2 = tk.Entry(info_frame2, width=10)
date_entry2.grid(row=0, column=2, padx=5)


# Time
tk.Label(info_frame2, text="Time").grid(
    row=0, column=3, padx=5, pady=5
)
time_entry2 = tk.Entry(info_frame2, width=10)
time_entry2.grid(row=0, column=4, padx=5)


# Price
tk.Label(info_frame2, text="Price").grid(
    row=0, column=5, padx=5, pady=5
)
price_entry2 = tk.Entry(info_frame2, width=10)
price_entry2.grid(row=0, column=6, padx=5)


# Submit button
submit_button = tk.Button(
    root,
    text="Submit",
    font=("Arial", 11, "bold"),
    width=15,
    command=submit_form
)
submit_button.pack(pady=10)


# Run application
root.mainloop()

