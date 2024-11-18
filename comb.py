import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from itertools import combinations_with_replacement, product
from collections import Counter
from typing import Dict, List
import json
import platform
import os
import pandas as pd
from PIL import Image, ImageTk, ImageSequence

class CombinationGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Combination Generator")
        self.root.geometry("1100x600")
        
        self.set_icon()
        
        self.categories = {}
        self.all_combinations = []
        self.combination_history = []
        
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.create_category_frame()
        self.create_items_frame()
        self.create_results_frame()
        
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(2, weight=1)
        
        style = ttk.Style()
        style.configure('Custom.TButton', padding=5)
        
        self.load_categories()
        
    def set_icon(self):
        if platform.system() == "Windows":
            self.root.iconbitmap("icon/sukuna.ico") 
        else:
            try:
                self.icon_image = ImageTk.PhotoImage(file="icon/sukuna.png") 
                self.root.iconphoto(True, self.icon_image)
            except Exception as e:
                print(f"Failed to set icon: {e}")

    def create_category_frame(self):
        category_frame = ttk.LabelFrame(self.main_frame, text="Add Category", padding="5")
        category_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(category_frame, text="Category Name:").grid(row=0, column=0, padx=5)
        self.category_entry = ttk.Entry(category_frame)
        self.category_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(category_frame, text="Quantity:").grid(row=0, column=2, padx=5)
        self.quantity_spinbox = ttk.Spinbox(category_frame, from_=1, to=100, width=5)
        self.quantity_spinbox.grid(row=0, column=3, padx=5)
        
        add_cat_btn = ttk.Button(category_frame, text="Add Category", 
                                command=self.add_category, style='Custom.TButton')
        add_cat_btn.grid(row=0, column=4, padx=5)

    def create_items_frame(self):
        items_frame = ttk.LabelFrame(self.main_frame, text="Categories and Items", padding="5")
        items_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(items_frame, text="Select Category:").grid(row=0, column=0, padx=5)
        self.category_combo = ttk.Combobox(items_frame, state="readonly")
        self.category_combo.grid(row=0, column=1, padx=5)
        
        ttk.Label(items_frame, text="Item Name:").grid(row=0, column=2, padx=5)
        self.item_entry = ttk.Entry(items_frame)
        self.item_entry.grid(row=0, column=3, padx=5)
        
        add_item_btn = ttk.Button(items_frame, text="Add Item", 
                                 command=self.add_item, style='Custom.TButton')
        add_item_btn.grid(row=0, column=4, padx=5)

    def create_results_frame(self):
        results_frame = ttk.LabelFrame(self.main_frame, text="Results", padding="5")
        results_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        self.category_text = tk.Text(results_frame, height=10, width=40)
        self.category_text.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        category_scroll = ttk.Scrollbar(results_frame, orient="vertical", 
                                      command=self.category_text.yview)
        category_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.category_text.configure(yscrollcommand=category_scroll.set)
        
        self.combinations_text = tk.Text(results_frame, height=10, width=40)
        self.combinations_text.grid(row=1, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        combinations_scroll = ttk.Scrollbar(results_frame, orient="vertical", 
                                          command=self.combinations_text.yview)
        combinations_scroll.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.combinations_text.configure(yscrollcommand=combinations_scroll.set)
        
        buttons_frame = ttk.Frame(results_frame)
        buttons_frame.grid(row=2, column=0, columnspan=2, pady=5)
        
        generate_btn = ttk.Button(buttons_frame, text="Generate Combinations", 
                                command=self.generate_combinations, style='Custom.TButton')
        generate_btn.grid(row=0, column=0, padx=5)
        
        self.export_excel_btn = ttk.Button(buttons_frame, text="Export to Excel", 
                                command=self.export_to_excel, style='Custom.TButton', state=tk.DISABLED)
        self.export_excel_btn.grid(row=0, column=1, padx=5)
        
        save_btn = ttk.Button(buttons_frame, text="Save Categories", 
                             command=self.save_categories, style='Custom.TButton')
        save_btn.grid(row=0, column=2, padx=5)
        
        clear_btn = ttk.Button(buttons_frame, text="Clear All", 
                              command=self.clear_all, style='Custom.TButton')
        clear_btn.grid(row=0, column=3, padx=5)
        
        self.gif_label = tk.Label(results_frame)
        self.gif_label.grid(row=0, column=2, rowspan=2, padx=10, pady=5, sticky=(tk.N, tk.S))

        self.credit_label = tk.Label(results_frame, text="made by @ralphmodales", font=("Arial", 10), fg="gray")
        self.credit_label.grid(row=2, column=2, padx=10, pady=(0, 10), sticky=(tk.S))

        self.display_gif("sukuna.gif")
        
        history_btn = ttk.Button(buttons_frame, text="View History", 
                         command=self.view_history, style='Custom.TButton')
        history_btn.grid(row=0, column=4, padx=5)

        export_history_btn = ttk.Button(buttons_frame, text="Export History", 
                                        command=self.export_history_to_excel, style='Custom.TButton')
        export_history_btn.grid(row=0, column=5, padx=5)
        
    def display_gif(self, gif_path):
        self.gif_image = Image.open(gif_path)
        self.gif_frames = [ImageTk.PhotoImage(frame) for frame in ImageSequence.Iterator(self.gif_image)]
        self.gif_index = 0
        
        def update_frame():
            if self.gif_frames:
                frame = self.gif_frames[self.gif_index]
                self.gif_label.config(image=frame)
                self.gif_index = (self.gif_index + 1) % len(self.gif_frames)
                self.root.after(100, update_frame) 
        
        update_frame()

    def add_category(self):
        category = self.category_entry.get().strip()
        try:
            quantity = int(self.quantity_spinbox.get())
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid quantity")
            return
            
        if category:
            if category not in self.categories:
                self.categories[category] = {"items": [], "quantity": quantity}
                self.update_category_combo()
                self.update_display()
                self.category_entry.delete(0, tk.END)
                self.quantity_spinbox.delete(0, tk.END)
                self.quantity_spinbox.insert(0, "1")
            else:
                messagebox.showwarning("Warning", "Category already exists!")
        else:
            messagebox.showwarning("Warning", "Please enter a category name!")

    def add_item(self):
        category = self.category_combo.get()
        item = self.item_entry.get().strip()
        
        if category and item:
            if item not in self.categories[category]["items"]:
                self.categories[category]["items"].append(item)
                self.update_display()
                self.item_entry.delete(0, tk.END)
            else:
                messagebox.showwarning("Warning", "Item already exists in this category!")
        else:
            messagebox.showwarning("Warning", "Please select a category and enter an item name!")

    def update_category_combo(self):
        self.category_combo['values'] = list(self.categories.keys())
        if self.categories:
            self.category_combo.set(list(self.categories.keys())[0])

    def update_display(self):
        self.category_text.delete(1.0, tk.END)
        for category, data in self.categories.items():
            self.category_text.insert(tk.END, f"\nCategory: {category} (Quantity: {data['quantity']})\n")
            for item in data["items"]:
                self.category_text.insert(tk.END, f"  - {item}\n")

    def generate_combinations(self):
        if not self.categories:
            messagebox.showwarning("Warning", "Please add categories and items first!")
            return
            
        for category, data in self.categories.items():
            if not data["items"]:
                messagebox.showwarning("Warning", f"Category '{category}' has no items!")
                return
        
        self.all_combinations = self.calculate_combinations()
        
        self.combinations_text.delete(1.0, tk.END)
        self.combinations_text.insert(tk.END, f"Total combinations: {len(self.all_combinations)}\n\n")
        
        for i, combo in enumerate(self.all_combinations[:10], 1):
            self.combinations_text.insert(tk.END, f"Combination {i}: ")
            combo_parts = []
            for category, data in self.categories.items():
                category_items = []
                for item in data['items']:
                    item_count = combo.get(item, 0)
                    if item_count > 0:
                        category_items.append(f"{item_count} pcs {item} {category}")
                if category_items:
                    combo_parts.append(" + ".join(category_items))
            
            self.combinations_text.insert(tk.END, " + ".join(combo_parts) + "\n")
        
        if len(self.all_combinations) > 10:
            self.combinations_text.insert(tk.END, f"... and {len(self.all_combinations) - 10} more combinations\n")
        
        self.export_excel_btn.config(state=tk.NORMAL)
        
        history_entry = {
        'timestamp': pd.Timestamp.now(),
        'categories': self.categories.copy(),
        'combinations': self.all_combinations
        }
        self.combination_history.append(history_entry)

    def view_history(self):
        if not self.combination_history:
            messagebox.showinfo("History", "No combination history available.")
            return

        history_window = tk.Toplevel(self.root)
        history_window.title("Combination Generation History")
        history_window.geometry("600x400")

        history_text = tk.Text(history_window, wrap=tk.WORD)
        history_text.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        for i, entry in enumerate(self.combination_history, 1):
            history_text.insert(tk.END, f"History Entry {i} - {entry['timestamp']}\n")
            history_text.insert(tk.END, "Categories:\n")
            for category, data in entry['categories'].items():
                history_text.insert(tk.END, f"  {category} (Quantity: {data['quantity']}): {', '.join(data['items'])}\n")
            history_text.insert(tk.END, f"Total Combinations: {len(entry['combinations'])}\n\n")

        scrollbar = ttk.Scrollbar(history_window, command=history_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        history_text.configure(yscrollcommand=scrollbar.set)

    def export_history_to_excel(self):
        if not self.combination_history:
            messagebox.showwarning("Warning", "No combination history available!")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            all_history_data = []
            for entry in self.combination_history:
                entry_data = {
                    'Timestamp': entry['timestamp'],
                    'Categories': json.dumps(entry['categories']),
                    'Total Combinations': len(entry['combinations'])
                }
                all_history_data.append(entry_data)
            
            df = pd.DataFrame(all_history_data)
            df.to_excel(file_path, index=False)
            
            messagebox.showinfo("Success", f"Combination history exported to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export history: {str(e)}")
            
    def calculate_combinations(self) -> List[Counter]:
        def generate_valid_combinations(category, items, total_quantity):
            unique_combos = []
            
            for distribution in combinations_with_replacement(range(len(items)), total_quantity):
                combo = Counter(items[i] for i in distribution)

                if combo not in unique_combos:
                    unique_combos.append(combo)
            
            return [Counter(combo) for combo in unique_combos]

        category_combinations = {}
        for category, data in self.categories.items():
            possible_items = data["items"]
            quantity = data["quantity"]
            
            category_combinations[category] = generate_valid_combinations(
                category, 
                possible_items, 
                quantity
            )
        
        categories = sorted(self.categories.keys())
        category_combos = [category_combinations[cat] for cat in categories]
        
        if len(category_combos) == 1:
            return category_combos[0]
        
        all_combinations = []
        for combo in product(*category_combos):
            combined_combo = Counter()
            for cat_combo in combo:
                combined_combo.update(cat_combo)
            all_combinations.append(combined_combo)
        
        return all_combinations
    
    def export_to_excel(self):
        if not self.all_combinations:
            messagebox.showwarning("Warning", "Generate combinations first!")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            headers = ["Combination"]
            
            excel_data = [headers]
            
            for i, combo in enumerate(self.all_combinations, 1):
                combo_parts = []
                
                for category, data in self.categories.items():
                    for item in data['items']:
                        item_count = combo.get(item, 0)
                        
                        if item_count > 0:
                            combo_parts.append(f"{item_count} pcs {item} {category}")
                
                combination_description = " + ".join(combo_parts)
                excel_data.append([combination_description])
            
            df = pd.DataFrame(excel_data[1:], columns=excel_data[0])
            df.to_excel(file_path, index=False)
            
            messagebox.showinfo("Success", f"Combinations exported to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export to Excel: {str(e)}")

    def save_categories(self):
        try:
            with open('categories.json', 'w') as f:
                json.dump(self.categories, f)
            messagebox.showinfo("Success", "Categories saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save categories: {str(e)}")

    def load_categories(self):
        try:
            if os.path.exists('categories.json'):
                with open('categories.json', 'r') as f:
                    self.categories = json.load(f)
                self.update_category_combo()
                self.update_display()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load categories: {str(e)}")

    def clear_all(self):
        if messagebox.askyesno("Confirm", "Are you sure you want to clear all categories and items?"):
            self.categories.clear()
            self.all_combinations.clear()
            self.category_text.delete(1.0, tk.END)
            self.combinations_text.delete(1.0, tk.END)
            self.update_category_combo()
            self.export_excel_btn.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = CombinationGeneratorGUI(root)
    root.mainloop()
