import os
import math
import multiprocessing
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

daily_cols = ['AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK']
degree_cols = ['AQ','AS','AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK']

def browse_file(entry_field):
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel and CSV files", "*.xlsx *.xls *.csv")]
    )
    if file_path:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, file_path)

def browse_folder(entry_field):
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, folder_path)

def parse_row_to_dict(row):
    row = list(row)
    # row is a list like [Player, Degrees, Degrees Count, Total, WinPercent]
    data = {}
    data['Player'] = row[0]
    data['Columns'] = [row[1][i:i+2] for i in range(0, len(row[1]), 2)]
    data['DegreesCount'] = row[2]
    data['WinPercent'] = row[4]
    return data

def multiprocess_rows(args):
    chunk, daily_df, raw_daily_df = args
    matches = []
    for idx, row in chunk:
        try:
            # print(f"Processing {idx} row...")
            row_dict = parse_row_to_dict(row)
            hist_row = pd.Series(row_dict)

            for i, daily_row in daily_df.iterrows():
                is_match = True
                daily_player = str(daily_row['Player'])
                hist_player = str(hist_row['Player'])

                if daily_player != hist_player:
                    is_match = False

                daily_degree_count = 0
                for col in hist_row['Columns']:
                    daily_degree_count += int(daily_row[col])

                if daily_degree_count != int(hist_row['DegreesCount']):
                    is_match = False
                
                if is_match:
                    print(f"***** Found the matching result of {idx} row *****")
                    matched_row = raw_daily_df.iloc[i].tolist()
                    cleaned_row = []
                    for val in list(row):
                        try:
                            cleaned_row.append(float(val))
                        except:
                            cleaned_row.append(val)
                    matched_row = matched_row + cleaned_row
                    matches.append(matched_row)
        except Exception as e:
            print(f"Skipping because of wrong format values as {e}")
            continue
    return matches

def process_files():
    try:
        daily_file = daily_entry.get()
        historical_folder = hist_entry.get()
        output_format = output_format_var.get()

        if not os.path.exists(daily_file) or not os.path.exists(historical_folder):
            messagebox.showerror("Error", "One or both files not found.")
            return

        raw_daily_df = pd.read_csv(daily_file, header=None) if daily_file.endswith('csv') else pd.read_excel(daily_file, header=None)

        # filter Daily Data
        columns_to_extract = [0] + list(range(41, 63))
        daily_df = raw_daily_df[columns_to_extract]

        column_labels = ['Player'] + daily_cols
        daily_df.columns = column_labels

        all_matches = []
        for file_name in os.listdir(historical_folder):
            input_path = os.path.join(historical_folder, file_name)
            raw_hist_df = pd.read_csv(input_path, header=None) if input_path.endswith('csv') else pd.read_excel(input_path, header=None)

            all_rows = list(raw_hist_df.iterrows())
            chunk_size = math.ceil(len(all_rows) / os.cpu_count())
            chunks = [all_rows[i:i + chunk_size] for i in range(0, len(all_rows), chunk_size)]

            args = [(chunk, daily_df, raw_daily_df) for chunk in chunks]
            with multiprocessing.get_context("spawn").Pool() as pool:
                one_result = pool.map(multiprocess_rows, args)

            flat_rows = [row for group in one_result for row in group if row]
            all_matches.extend(flat_rows)
            print(f"--------------- {file_name} Processed succesfully ---------------")

        if all_matches:
            matched_df = pd.DataFrame(all_matches)
            matched_df = matched_df.dropna(how='all')
            if output_format == 'excel':
                output_path = os.path.splitext(daily_file)[0] + '_Matches.xlsx'
                matched_df.to_excel(output_path, index=False, header=None)
            else:
                output_path = os.path.splitext(daily_file)[0] + '_Matches.csv'
                matched_df.to_csv(output_path, index=False, header=None)
            # Extend your processing logic here
            messagebox.showinfo("Success", f"Processing finished...")
        else:
            messagebox.showerror("Error", f"NO Matches found...")
    except Exception as e:
        messagebox.showerror("Error", f"Error occured as {e}")

if __name__ == "__main__":
    # GUI setup
    root = tk.Tk()
    root.title("Matcher Processor with non-coloring")

    # Row 1 - Daily Input
    tk.Label(root, text="Daily File:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
    daily_entry = tk.Entry(root, width=50)
    daily_entry.grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=lambda: browse_file(daily_entry)).grid(row=0, column=2, padx=5, pady=5)

    # Row 2 - Historical Input
    tk.Label(root, text="Historical % Input File:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
    hist_entry = tk.Entry(root, width=50)
    hist_entry.grid(row=1, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=lambda: browse_folder(hist_entry)).grid(row=1, column=2, padx=5, pady=5)

    # Row 3 - Excel or CSV Option
    output_format_var = tk.StringVar(value="csv")  # Default to CSV

    tk.Label(root, text="Output Format:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
    radio_frame = tk.Frame(root)
    radio_frame.grid(row=2, column=1, padx=5, pady=5, sticky="w")

    tk.Radiobutton(radio_frame, text="CSV", variable=output_format_var, value="csv").pack(side="left")
    tk.Radiobutton(radio_frame, text="Excel", variable=output_format_var, value="excel").pack(side="left")

    # Row 4 - Process Button
    tk.Button(root, text="Process", command=process_files, width=20).grid(row=3, column=1, pady=15)

    root.mainloop()