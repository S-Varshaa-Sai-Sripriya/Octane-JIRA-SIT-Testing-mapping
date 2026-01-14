import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from openpyxl import Workbook
import threading

class OctaneJiraMapperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Octane ID - JIRA ID Mapper")
        self.root.geometry("750x650")
        self.root.resizable(True, True)
        
        self.input_file = None
        self.output_df = None
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Title
        title_label = ttk.Label(root, text="Octane ID - JIRA ID Mapper", font=("Arial", 16, "bold"))
        title_label.pack(pady=20)
        
        # File selection frame
        file_frame = ttk.LabelFrame(root, text="Step 1: Select Input File", padding=10)
        file_frame.pack(padx=20, pady=10, fill="x")
        
        self.file_label = ttk.Label(file_frame, text="No file selected", foreground="gray")
        self.file_label.pack(side="left", padx=10)
        
        browse_btn = ttk.Button(file_frame, text="Browse", command=self.browse_file)
        browse_btn.pack(side="right", padx=10)
        
        # Action frame
        action_frame = ttk.LabelFrame(root, text="Step 2: Process Data", padding=10)
        action_frame.pack(padx=20, pady=10, fill="x")
        
        compute_btn = ttk.Button(action_frame, text="Compute Mapping", command=self.compute_mapping)
        compute_btn.pack(side="left", padx=10)
        
        # Progress bar
        self.progress = ttk.Progressbar(action_frame, mode='indeterminate')
        self.progress.pack(side="right", padx=10, fill="x", expand=True)
        
        # Output frame
        output_frame = ttk.LabelFrame(root, text="Preview (First 10 rows)", padding=10)
        output_frame.pack(padx=20, pady=10, fill="both", expand=True)
        
        # Treeview for preview
        columns = ("Test Team", "Octane ID", "JIRA ID")
        self.tree = ttk.Treeview(output_frame, columns=columns, height=12, show='headings')
        
        for col in columns:
            self.tree.column(col, width=200, anchor='w')
            self.tree.heading(col, text=col)
        
        scrollbar = ttk.Scrollbar(output_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Status and Download frame
        bottom_frame = ttk.Frame(root)
        bottom_frame.pack(pady=10, fill="x", padx=20)
        
        status_label = ttk.Label(bottom_frame, text="Ready to process files", foreground="blue")
        status_label.pack(pady=5)
        self.status_label = status_label
        
        # Download button - large and centered
        self.download_btn = ttk.Button(bottom_frame, text="📥 Download Output Excel", 
                                        command=self.save_output, state="disabled")
        self.download_btn.pack(pady=10, ipadx=20, ipady=5)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.input_file = file_path
            self.file_label.config(text=f"...{file_path[-50:]}", foreground="black")
            self.status_label.config(text=f"File loaded: {file_path.split('/')[-1]}", foreground="blue")
    
    def compute_mapping(self):
        if not self.input_file:
            messagebox.showerror("Error", "Please select an input file first!")
            return
        
        # Process directly (simpler approach)
        self._process_file()
    
    def _process_file(self):
        try:
            self.progress.start()
            self.status_label.config(text="Processing...", foreground="orange")
            self.download_btn.config(state="disabled")
            self.root.update()
            
            # Read input file
            df = pd.read_excel(self.input_file)
            
            # Create output data
            output_rows = []
            
            for idx, row in df.iterrows():
                test_team = row.get('Test Team', '')
                octane_id = row.get('ID', '')
                jira_ids = row.get('Test: JIRA ID', '')
                
                # Skip rows with missing critical data
                if pd.isna(test_team) or pd.isna(octane_id):
                    continue
                
                # Convert to string
                test_team = str(test_team).strip()
                octane_id = str(int(octane_id)) if pd.notna(octane_id) else ''
                jira_ids_str = str(jira_ids).strip() if pd.notna(jira_ids) else ''
                
                if not jira_ids_str or jira_ids_str == '':
                    # Add row even if no JIRA ID
                    output_rows.append({
                        'Test Team': test_team,
                        'Octane ID': octane_id,
                        'JIRA ID': ''
                    })
                else:
                    # Split JIRA IDs and create a row for each
                    jira_list = [j.strip() for j in jira_ids_str.split(',')]
                    for jira_id in jira_list:
                        output_rows.append({
                            'Test Team': test_team,
                            'Octane ID': octane_id,
                            'JIRA ID': jira_id
                        })
            
            # Create output dataframe
            self.output_df = pd.DataFrame(output_rows)
            
            # Update preview
            self._update_preview()
            
            self.progress.stop()
            self.download_btn.config(state="normal")
            self.status_label.config(
                text=f"✓ Success! Processed {len(self.output_df)} rows", 
                foreground="green"
            )
            messagebox.showinfo("Success", f"Successfully processed {len(self.output_df)} rows!\n\nClick 'Download Output Excel' to save the file.")
            
        except Exception as e:
            self.progress.stop()
            self.status_label.config(text="Error occurred!", foreground="red")
            messagebox.showerror("Error", f"Error processing file:\n{str(e)}")
    
    def _update_preview(self):
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add first 10 rows
        for idx, row in self.output_df.head(10).iterrows():
            values = (row['Test Team'], row['Octane ID'], row['JIRA ID'])
            self.tree.insert('', 'end', values=values)
    
    def save_output(self):
        if self.output_df is None or self.output_df.empty:
            messagebox.showerror("Error", "No data to save!")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="octane_jira_mapping_output.xlsx"
        )
        
        if file_path:
            try:
                self.output_df.to_excel(file_path, sheet_name='Mapped Data', index=False)
                messagebox.showinfo("Success", f"File saved successfully!\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Error saving file:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = OctaneJiraMapperApp(root)
    root.mainloop()
