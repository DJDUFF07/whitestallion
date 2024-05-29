import pandas as pd
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox

# Initialize the main application window with customtkinter
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

root = ctk.CTk()
root.state('zoomed')
root.configure(bg='black')

# Define colors
colour1 = '#0a0b0c'
colour2 = '#0000cd'
colour3 = '#1e90ff'
colour4 = 'black'

class ExcelCombinerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Combiner")
        self.files = []
        self.dataframes = []
        self.headings_vars = []
        self.selected_headings = {}
        self.reference_heading = None

        self.add_files_button = ctk.CTkButton(
            root,
            fg_color=colour2,
            text_color='white',
            width=200,
            height=50,
            text="STEP 1. ADD EXCEL FILES",
            font=('Arial', 15, 'bold'),
            command=self.load_excel_files
        )
        self.add_files_button.pack(pady=10)

        self.ref_heading_button = ctk.CTkButton(
            root,
            fg_color=colour2,
            text_color='white',
            width=300,
            height=50,
            text="STEP 2. SELECT REFERENCE HEADING",
            font=('Arial', 15, 'bold'),
            command=self.select_reference_heading,
            state=ctk.DISABLED
        )
        self.ref_heading_button.pack(pady=10)

        self.drag_drop_frame = ctk.CTkFrame(root, fg_color=colour1)
        self.drag_drop_frame.pack(fill=ctk.BOTH, expand=True, pady=20)

        self.headings_frame = ctk.CTkFrame(self.drag_drop_frame, fg_color=colour1)
        self.headings_frame.pack(side=ctk.LEFT, fill=ctk.BOTH, expand=True, padx=10, pady=10)

        self.drop_zone_frame = ctk.CTkFrame(self.drag_drop_frame, fg_color=colour1)
        self.drop_zone_frame.pack(side=ctk.RIGHT, fill=ctk.BOTH, expand=True, padx=10, pady=10)

        self.combine_files_button = ctk.CTkButton(
            root,
            fg_color=colour2,
            text_color='white',
            width=350,
            height=50,
            text="STEP 3. COMBINE SELECTED HEADINGS",
            font=('Arial', 15, 'bold'),
            command=self.create_combined_excel,
            state=ctk.DISABLED
        )
        self.combine_files_button.pack(pady=10)

        self.canvas = tk.Canvas(self.headings_frame, bg=colour1)
        self.scrollbar = ctk.CTkScrollbar(self.headings_frame, orientation="vertical", command=self.canvas.yview)
        self.scrollable_frame = ctk.CTkFrame(self.canvas, fg_color=colour1)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.drop_zone = tk.Listbox(self.drop_zone_frame, bg=colour1, fg='white', selectmode=tk.SINGLE)
        self.drop_zone.pack(fill=ctk.BOTH, expand=True)

    def load_excel_files(self):
        file_paths = filedialog.askopenfilenames(title='Select Excel Files', filetypes=[('Excel files', '*.xlsx *.xls')])
        if file_paths:
            self.files = file_paths
            self.dataframes = [pd.read_excel(file) for file in self.files]
            self.display_headings()
            messagebox.showinfo("Files Loaded", f"Loaded {len(self.files)} files.")
            self.ref_heading_button.configure(state=ctk.NORMAL)
            self.combine_files_button.configure(state=ctk.NORMAL)

    def display_headings(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        self.headings_vars = []
        for file, df in zip(self.files, self.dataframes):
            ctk.CTkLabel(self.scrollable_frame, text=f"{file}", fg_color=colour1, text_color='white').pack(anchor="w")
            for heading in df.columns:
                lbl = ctk.CTkLabel(self.scrollable_frame, text=heading, fg_color=colour3, text_color='white')
                lbl.pack(anchor="w", fill=ctk.X, padx=5, pady=2)
                lbl.bind("<ButtonPress-1>", self.on_drag_start)
                self.headings_vars.append(lbl)

    def on_drag_start(self, event):
        event.widget.startX = event.x
        event.widget.startY = event.y
        event.widget.bind("<B1-Motion>", self.on_drag_motion)
        event.widget.bind("<ButtonRelease-1>", self.on_drop)

    def on_drag_motion(self, event):
        x = event.widget.winfo_x() - event.widget.startX + event.x
        y = event.widget.winfo_y() - event.widget.startY + event.y
        event.widget.place(x=x, y=y)

    def on_drop(self, event):
        widget = event.widget
        widget.unbind("<B1-Motion>")
        widget.unbind("<ButtonRelease-1>")
        x, y = widget.winfo_x(), widget.winfo_y()
        widget.place_forget()

        if self.drop_zone_frame.winfo_containing(x, y):
            self.drop_zone.insert(ctk.END, widget.cget("text"))
            widget.destroy()

    def select_reference_heading(self):
        selected = self.drop_zone.curselection()
        if selected:
            self.reference_heading = self.drop_zone.get(selected[0])
            messagebox.showinfo("Reference Heading Selected", f"'{self.reference_heading}' selected as the reference heading.")
            self.drop_zone.itemconfig(selected[0], {'bg': '#ff0000'})
        else:
            messagebox.showwarning("No Selection", "Please select a heading to be the reference.")

    def create_combined_excel(self):
        if not self.reference_heading:
            messagebox.showwarning("No Reference Heading", "Please select a reference heading first.")
            return

        combined_data = pd.DataFrame()
        ref_data = pd.DataFrame()

        for df in self.dataframes:
            if self.reference_heading in df.columns:
                ref_data = df[[self.reference_heading]].drop_duplicates().sort_values(by=self.reference_heading)
                break

        if ref_data.empty:
            messagebox.showwarning("Reference Heading Missing", "Reference heading not found in any file.")
            return

        combined_data[self.reference_heading] = ref_data[self.reference_heading]

        selected_headings = self.drop_zone.get(0, tk.END)
        for df in self.dataframes:
            selected_columns = [heading for heading in selected_headings if heading in df.columns and heading != self.reference_heading]
            if selected_columns:
                df = df[[self.reference_heading] + selected_columns]
                combined_data = combined_data.merge(df, on=self.reference_heading, how='left')

        combined_data = combined_data.drop_duplicates(subset=[self.reference_heading])

        for df in self.dataframes:
            selected_columns = [heading for heading in selected_headings if heading in df.columns and heading != self.reference_heading]
            if selected_columns:
                df = df[[self.reference_heading] + selected_columns]
                df = df[~df[self.reference_heading].isin(combined_data[self.reference_heading])]
                combined_data = pd.concat([combined_data, df], axis=0)

        if not combined_data.empty:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")])
            if save_path:
                combined_data.to_excel(save_path, index=False)
                messagebox.showinfo("File Saved", f"Combined Excel file saved as {save_path}")
        else:
            messagebox.showwarning("No Data", "No columns were selected for combining.")

if __name__ == "__main__":
    app = ExcelCombinerApp(root)

root.mainloop()
