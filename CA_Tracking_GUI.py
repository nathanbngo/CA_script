"""
Corporate Action Tracking System - GUI
Simple interface for viewing CA tracking summary and running updates.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import os
import subprocess
import sys
from datetime import datetime
import pandas as pd
from pathlib import Path

# Configuration (same as main script)
OUTPUT_FILE = r"F:\Trade Support\Corporate Actions\CA check\CA Raw file\CA_Tracking.xlsx"
SCRIPT_FILE = "CA_Tracking_System.py"

class CATrackingGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Corporate Action Tracking System")
        self.root.geometry("600x500")
        self.root.resizable(False, False)
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        self.setup_ui()
        self.refresh_stats()
        
        # Auto-refresh every 30 seconds
        self.auto_refresh()
    
    def setup_ui(self):
        """Set up the user interface"""
        # Header
        header_frame = tk.Frame(self.root, bg="#366092", height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="Corporate Action Tracking",
            font=("Arial", 18, "bold"),
            bg="#366092",
            fg="white"
        )
        title_label.pack(pady=15)
        
        # Main content frame
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Stats section
        stats_frame = tk.LabelFrame(main_frame, text="Current Status", font=("Arial", 12, "bold"), padx=15, pady=15)
        stats_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.tab1_label = tk.Label(stats_frame, text="Next 15 Days: Loading...", font=("Arial", 11))
        self.tab1_label.pack(anchor=tk.W, pady=5)
        
        self.tab2_label = tk.Label(stats_frame, text="Last 7 Days: Loading...", font=("Arial", 11))
        self.tab2_label.pack(anchor=tk.W, pady=5)
        
        self.archive_label = tk.Label(stats_frame, text="Archive: Loading...", font=("Arial", 11))
        self.archive_label.pack(anchor=tk.W, pady=5)
        
        self.last_update_label = tk.Label(stats_frame, text="Last Update: Never", font=("Arial", 10), fg="gray")
        self.last_update_label.pack(anchor=tk.W, pady=5)
        
        # Urgent items section
        urgent_frame = tk.LabelFrame(main_frame, text="Urgent Items (< 3 days)", font=("Arial", 12, "bold"), padx=15, pady=15)
        urgent_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Scrollable list for urgent items
        scrollbar = tk.Scrollbar(urgent_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.urgent_listbox = tk.Listbox(
            urgent_frame,
            font=("Arial", 10),
            yscrollcommand=scrollbar.set,
            height=8
        )
        self.urgent_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.urgent_listbox.yview)
        
        # Buttons frame
        buttons_frame = tk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X)
        
        # Update button
        self.update_button = tk.Button(
            buttons_frame,
            text="Update CA Tracking",
            font=("Arial", 12, "bold"),
            bg="#4CAF50",
            fg="white",
            padx=20,
            pady=10,
            command=self.run_update,
            cursor="hand2"
        )
        self.update_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Open Excel button
        self.open_excel_button = tk.Button(
            buttons_frame,
            text="Open Excel File",
            font=("Arial", 12),
            bg="#2196F3",
            fg="white",
            padx=20,
            pady=10,
            command=self.open_excel,
            cursor="hand2"
        )
        self.open_excel_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Refresh button
        self.refresh_button = tk.Button(
            buttons_frame,
            text="Refresh Stats",
            font=("Arial", 12),
            bg="#FF9800",
            fg="white",
            padx=20,
            pady=10,
            command=self.refresh_stats,
            cursor="hand2"
        )
        self.refresh_button.pack(side=tk.LEFT)
        
        # Status bar
        self.status_label = tk.Label(
            self.root,
            text="Ready",
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W,
            padx=10,
            pady=5,
            bg="#f0f0f0"
        )
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
    
    def refresh_stats(self):
        """Refresh statistics from Excel file"""
        try:
            if not os.path.exists(OUTPUT_FILE):
                self.tab1_label.config(text="Next 15 Days: No data (file not found)")
                self.tab2_label.config(text="Last 7 Days: No data (file not found)")
                self.archive_label.config(text="Archive: No data (file not found)")
                self.last_update_label.config(text="Last Update: File not found")
                self.urgent_listbox.delete(0, tk.END)
                self.urgent_listbox.insert(0, "No data available")
                return
            
            # Get file modification time
            mod_time = os.path.getmtime(OUTPUT_FILE)
            mod_datetime = datetime.fromtimestamp(mod_time)
            self.last_update_label.config(
                text=f"Last Update: {mod_datetime.strftime('%Y-%m-%d %H:%M:%S')}"
            )
            
            # Load data from Excel
            try:
                tab1_df = pd.read_excel(OUTPUT_FILE, sheet_name="Next 15 Days")
                tab1_count = len(tab1_df)
            except:
                tab1_count = 0
                tab1_df = pd.DataFrame()
            
            try:
                tab2_df = pd.read_excel(OUTPUT_FILE, sheet_name="Last 7 Days")
                tab2_count = len(tab2_df)
            except:
                tab2_count = 0
                tab2_df = pd.DataFrame()
            
            try:
                archive_df = pd.read_excel(OUTPUT_FILE, sheet_name="Archive")
                archive_count = len(archive_df)
            except:
                archive_count = 0
                archive_df = pd.DataFrame()
            
            # Update labels
            self.tab1_label.config(text=f"Next 15 Days: {tab1_count} CAs")
            self.tab2_label.config(text=f"Last 7 Days: {tab2_count} CAs")
            self.archive_label.config(text=f"Archive: {archive_count} CAs")
            
            # Get urgent items (deadline < 3 days)
            self.update_urgent_items(tab1_df)
            
            self.status_label.config(text="Stats refreshed successfully")
            
        except Exception as e:
            self.status_label.config(text=f"Error refreshing stats: {str(e)}")
            messagebox.showerror("Error", f"Failed to refresh stats:\n{str(e)}")
    
    def update_urgent_items(self, tab1_df):
        """Update urgent items list"""
        self.urgent_listbox.delete(0, tk.END)
        
        if tab1_df.empty or 'Deadline Date' not in tab1_df.columns:
            self.urgent_listbox.insert(0, "No urgent items")
            return
        
        try:
            from datetime import date, timedelta
            today = date.today()
            urgent_date = today + timedelta(days=3)
            
            urgent_count = 0
            for idx, row in tab1_df.iterrows():
                deadline_str = str(row.get('Deadline Date', ''))
                if not deadline_str or deadline_str == '':
                    continue
                
                try:
                    # Parse date (could be string or datetime)
                    if isinstance(deadline_str, str):
                        deadline = pd.to_datetime(deadline_str).date()
                    else:
                        deadline = deadline_str.date() if hasattr(deadline_str, 'date') else deadline_str
                    
                    if deadline <= urgent_date:
                        security_name = str(row.get('Security Name', 'Unknown'))[:40]
                        event_type = str(row.get('Event Type', 'Unknown'))[:30]
                        days_left = (deadline - today).days
                        
                        item_text = f"{security_name} | {event_type} | {days_left} day(s) left"
                        self.urgent_listbox.insert(tk.END, item_text)
                        urgent_count += 1
                        
                        if urgent_count >= 20:  # Limit to 20 items
                            break
                except:
                    continue
            
            if urgent_count == 0:
                self.urgent_listbox.insert(0, "No urgent items (< 3 days)")
                
        except Exception as e:
            self.urgent_listbox.insert(0, f"Error loading urgent items: {str(e)}")
    
    def run_update(self):
        """Run the CA tracking update script"""
        self.update_button.config(state=tk.DISABLED, text="Updating...")
        self.status_label.config(text="Running update... Please wait...")
        self.root.update()
        
        try:
            # Import and run the main script directly
            # This works when running as .exe (bundled) or as Python script
            if getattr(sys, 'frozen', False):
                # Running as compiled executable - script is bundled
                script_path = os.path.join(sys._MEIPASS, SCRIPT_FILE)
            else:
                # Running as Python script
                script_path = os.path.join(os.path.dirname(__file__), SCRIPT_FILE)
            
            if not os.path.exists(script_path):
                raise FileNotFoundError(f"Script not found: {script_path}")
            
            # Import the module dynamically
            import importlib.util
            spec = importlib.util.spec_from_file_location("ca_tracking_system", script_path)
            if spec and spec.loader:
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                # Run the main function (suppress the GUI message box, we'll show our own)
                module.main()
                self.status_label.config(text="Update completed successfully!")
                messagebox.showinfo("Success", "CA Tracking updated successfully!")
                self.refresh_stats()
            else:
                raise ImportError("Could not load script module")
                
        except Exception as e:
            self.status_label.config(text="Update failed")
            error_msg = str(e)
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to run update:\n{error_msg}")
        finally:
            self.update_button.config(state=tk.NORMAL, text="Update CA Tracking")
    
    def open_excel(self):
        """Open the Excel file"""
        try:
            if not os.path.exists(OUTPUT_FILE):
                messagebox.showerror("Error", f"Excel file not found:\n{OUTPUT_FILE}")
                return
            
            # Open with default application
            if sys.platform == "win32":
                os.startfile(OUTPUT_FILE)
            elif sys.platform == "darwin":
                subprocess.run(["open", OUTPUT_FILE])
            else:
                subprocess.run(["xdg-open", OUTPUT_FILE])
            
            self.status_label.config(text="Opening Excel file...")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel file:\n{str(e)}")
    
    def auto_refresh(self):
        """Auto-refresh stats every 30 seconds"""
        self.refresh_stats()
        self.root.after(30000, self.auto_refresh)  # 30 seconds


def main():
    """Main entry point"""
    root = tk.Tk()
    app = CATrackingGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

