import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

class DataAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Corporate Data Analyzer")
        self.root.geometry("1100x700")

        self.file_path = None
        self.df = None
        self.report_df = None
        self.canvas = None

        self.create_widgets()

    def create_widgets(self):
        # ===== FILE SELECTION =====
        frame_file = tk.Frame(self.root)
        frame_file.pack(pady=10)

        tk.Button(frame_file, text="Browse File", command=self.browse_file).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_file, text="Read File", command=self.read_file).pack(side=tk.LEFT, padx=5)

        self.file_label = tk.Label(self.root, text="No file selected")
        self.file_label.pack()

        # ===== DATA INFO =====
        self.info_label = tk.Label(self.root, text="")
        self.info_label.pack(pady=10)

        # ===== REPORT BUILDER =====
        frame_report = tk.Frame(self.root)
        frame_report.pack(pady=10)

        tk.Label(frame_report, text="Group By").grid(row=0, column=0)
        self.group_col = ttk.Combobox(frame_report)
        self.group_col.grid(row=0, column=1)

        tk.Label(frame_report, text="Aggregation").grid(row=0, column=2)
        self.agg_method = ttk.Combobox(frame_report, values=["sum", "mean", "max", "min", "count", "median"])
        self.agg_method.grid(row=0, column=3)

        tk.Label(frame_report, text="Value Column").grid(row=0, column=4)
        self.value_col = ttk.Combobox(frame_report)
        self.value_col.grid(row=0, column=5)

        tk.Button(self.root, text="Preview Report", command=self.generate_report).pack(pady=5)
        tk.Button(self.root, text="Export Report", command=self.export_report).pack(pady=5)

        # ===== TABLE DISPLAY =====
        self.tree = ttk.Treeview(self.root)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # ===== CHART BUILDER =====
        frame_chart = tk.Frame(self.root)
        frame_chart.pack(pady=10)

        tk.Label(frame_chart, text="Chart Type").grid(row=0, column=0)
        self.chart_type = ttk.Combobox(frame_chart, values=["Bar", "Column", "Line", "Pie"])
        self.chart_type.grid(row=0, column=1)

        tk.Button(frame_chart, text="Preview Chart", command=self.plot_chart).grid(row=0, column=2, padx=5)
        tk.Button(frame_chart, text="Export Chart", command=self.export_chart).grid(row=0, column=3, padx=5)

    # ===== FILE HANDLING =====
    def browse_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx")])
        if self.file_path:
            self.file_label.config(text=self.file_path)

    def read_file(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select a file first.")
            return

        try:
            if self.file_path.endswith(".csv"):
                self.df = pd.read_csv(self.file_path)
            else:
                self.df = pd.read_excel(self.file_path)

            rows, cols = self.df.shape
            columns = list(self.df.columns)

            self.info_label.config(text=f"Rows: {rows} | Columns: {cols}\n{columns}")

            self.detect_columns()

        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ===== COLUMN DETECTION =====
    def detect_columns(self):
        text_cols = self.df.select_dtypes(include='object').columns.tolist()
        num_cols = self.df.select_dtypes(include='number').columns.tolist()

        # Try converting numeric-like text
        for col in text_cols[:]:
            try:
                self.df[col] = pd.to_numeric(self.df[col])
                num_cols.append(col)
                text_cols.remove(col)
            except:
                pass

        self.group_col['values'] = text_cols
        self.value_col['values'] = num_cols

    # ===== REPORT GENERATION =====
    def generate_report(self):
        if self.df is None:
            messagebox.showerror("Error", "Please read the file first.")
            return

        group = self.group_col.get()
        agg = self.agg_method.get()
        value = self.value_col.get()

        if not group or not agg or not value:
            messagebox.showerror("Error", "Please select all fields.")
            return

        try:
            self.report_df = self.df.groupby(group)[value].agg(agg).reset_index()
            self.report_df = self.report_df.sort_values(by=value, ascending=False)

            self.display_table(self.report_df)

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def display_table(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"

        for col in df.columns:
            self.tree.heading(col, text=col)

        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    # ===== EXPORT REPORT =====
    def export_report(self):
        if self.report_df is None:
            messagebox.showerror("Error", "No report to export.")
            return

        folder = os.path.dirname(self.file_path)
        save_path = filedialog.asksaveasfilename(initialdir=folder,
                                                 defaultextension=".xlsx",
                                                 filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])

        if save_path:
            if save_path.endswith(".csv"):
                self.report_df.to_csv(save_path, index=False)
            else:
                self.report_df.to_excel(save_path, index=False)

            messagebox.showinfo("Success", "Report exported successfully.")

    # ===== CHART =====
    def plot_chart(self):
        if self.report_df is None:
            messagebox.showerror("Error", "Generate report first.")
            return

        chart = self.chart_type.get()
        if not chart:
            messagebox.showerror("Error", "Select chart type.")
            return

        if self.canvas:
            self.canvas.get_tk_widget().destroy()

        fig, ax = plt.subplots()

        x = self.report_df.iloc[:, 0]
        y = self.report_df.iloc[:, 1]

        if chart == "Bar" or chart == "Column":
            ax.bar(x, y)
        elif chart == "Line":
            ax.plot(x, y)
        elif chart == "Pie":
            ax.pie(y, labels=x, autopct='%1.1f%%')

        ax.set_title("Chart Preview")

        self.canvas = FigureCanvasTkAgg(fig, master=self.root)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack()

    # ===== EXPORT CHART =====
    def export_chart(self):
        if self.report_df is None:
            messagebox.showerror("Error", "Generate report first.")
            return

        folder = os.path.dirname(self.file_path)
        save_path = os.path.join(folder, "chart.png")

        plt.savefig(save_path)
        messagebox.showinfo("Success", f"Chart saved at {save_path}")


# ===== RUN APP =====
if __name__ == "__main__":
    root = tk.Tk()
    app = DataAnalyzerApp(root)
    root.mainloop()