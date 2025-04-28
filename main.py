import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from asammdf import MDF
from docx import Document
from docx.shared import Inches
import io
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox

def _get_time_range(signals: dict) -> np.ndarray:
    # Function to determine the start and end time of the measurement
    if signals is None:
        return np.empty(shape=(0,))
    timeStart = 1e10
    timeEnd = 0
    for _, signal in signals.items():
        if signal.timestamps is None:
            continue
        try:
            timeEnd = max(timeEnd, signal.timestamps[-1])
            timeStart = min(timeStart, signal.timestamps[0])
        except IndexError as ie:
            print(ie)
    timeVector = np.arange(round(timeStart, 3), round(timeEnd, 3), 0.01)
    return timeVector

def _convert_signals(signals: dict, convert: bool):
    # Function to convert ASAM MDF signals to a unified time vector
    time_vector = _get_time_range(signals)
    if convert:
        df = pd.DataFrame(time_vector, columns=['time'])
        for key, signal in signals.items():
            syncSignal = signal.interp(time_vector, 0, 0)
            dataSignal = list(zip(syncSignal.timestamps, syncSignal.samples))
            dfSignal = pd.DataFrame(dataSignal, columns=['time', key])
            df = pd.merge(df, dfSignal, on='time')
        return df
    else:
        for key, signal in signals.items():
            syncSignal = signal.interp(time_vector, 0, 0)
            signal.samples = syncSignal.samples
        signals['time'] = time_vector
        return signals

def process_data(filename, convert=False):
    # Function to process the MDF4 file and extract the required signals
    try:
        measurement = MDF(filename)
    except Exception as e:
        print(f"Error opening the file {filename}: {e}")
        return None
    
    signalsInMeas = measurement.channels_db
    
    channel_list = [
        'DcrInEgoM_psid_Act',
        'DcrInEgoM_agwFA_Ste',
        'QU_FN_FDR'
    ]
    
    mappedSignals = {}
    for signal in channel_list:
        for sub, indexgroup in signalsInMeas.items():
            if sub == signal:
                indexFindings = indexgroup[0]
                group = indexFindings[0]
                index = indexFindings[1]
                mappedSignals[signal] = measurement.get(signal,
                                                        group=group,
                                                        index=index)
    convertedSignals = _convert_signals(mappedSignals, convert)
    print("Processing data from file:", filename)
    return convertedSignals

def process_mf4_data(mf4_data, filename, progress_listbox):
    # Function to process the MF4 data and perform analysis
    if mf4_data is None or mf4_data.empty:
        progress_listbox.insert(tk.END, f"MF4 data is empty or invalid for file: {filename}")
        return None

    # Find the first non-512 value from the end, then where it is 512 again
    index_qu_fn_fdr = None
    found_non_512 = False
    for i in range(len(mf4_data["QU_FN_FDR"]) - 1, -1, -1):
        value = mf4_data["QU_FN_FDR"].iloc[i]
        if not found_non_512 and value != 512:
            found_non_512 = True
        elif found_non_512 and value == 512:
            index_qu_fn_fdr = i
            break

    if index_qu_fn_fdr is None:
        progress_listbox.insert(tk.END, f"Could not find the required transition in QU_FN_FDR for file: {filename}")
        return None

    index_dcr_in_ego = 2
    for i in range(index_qu_fn_fdr - 30, index_qu_fn_fdr):
        if i > 0 and abs(mf4_data["DcrInEgoM_agwFA_Ste"].iloc[i] - mf4_data["DcrInEgoM_agwFA_Ste"].iloc[i - 1]) > 0.009:
            index_dcr_in_ego = i - 2
            break

    start_index = min(index_dcr_in_ego, index_dcr_in_ego)
    mf4_data_reduced = mf4_data.iloc[start_index:].reset_index(drop=True)

    psi_peak_index = mf4_data_reduced["DcrInEgoM_psid_Act"].abs().idxmax()
    psi_peak = mf4_data_reduced["DcrInEgoM_psid_Act"].iloc[psi_peak_index]

    T_0 = None
    for i in range(psi_peak_index, len(mf4_data_reduced["DcrInEgoM_agwFA_Ste"])):
        if abs(mf4_data_reduced["DcrInEgoM_agwFA_Ste"].iloc[i] - mf4_data_reduced["DcrInEgoM_agwFA_Ste"].iloc[-1]) < 0.005:
            T_0 = mf4_data_reduced["time"].iloc[i]
            break
    if T_0 is None:
        progress_listbox.insert(tk.END, f"T_0 could not be determined for file: {filename}")
        return

    tStartSine = mf4_data_reduced["time"].iloc[0]
    time = mf4_data_reduced["time"] - tStartSine
    T_0 = T_0 - tStartSine
    plt.figure(figsize=(10, 6))
    for column in mf4_data_reduced.columns:
        if column != "time" and column != "QU_FN_FDR":
            plt.plot(time, mf4_data_reduced[column], label=column)
    
    if T_0 is not None:
        plt.axvline(x=T_0, color='red', linestyle='--', label=f'T_0={T_0:.3f}s')
    psi_peak_time = mf4_data_reduced["time"].iloc[
        mf4_data_reduced["DcrInEgoM_psid_Act"].abs().idxmax()
    ] - tStartSine
    plt.axvline(x=psi_peak_time, color='green', linestyle='--', label=f'Psid Peak (Time={psi_peak_time:.3f}s, Val={psi_peak:.3f})')

    psi_at_t0_plus_1 = mf4_data_reduced["DcrInEgoM_psid_Act"].iloc[
        (mf4_data_reduced["time"] - tStartSine).sub(T_0 + 1).abs().idxmin()
    ]
    plt.axvline(x=T_0+1, color='black', linestyle='--', label=f'Psid(T_0+1)={psi_at_t0_plus_1:.3f}')
    plt.fill_betweenx(
        [-psi_peak * 0.35, psi_peak * 0.35],
        T_0 + 0.95,
        T_0 + 1.05,
        color='green',
        alpha=0.3,
        label='35% Psid Peak at T_0+1'
    )

    within_35_percent = (abs(psi_at_t0_plus_1) <= abs(psi_peak) * 0.35)

    psi_at_t0_plus_1p75 = mf4_data_reduced["DcrInEgoM_psid_Act"].iloc[
        (mf4_data_reduced["time"] - tStartSine).sub(T_0 + 1.75).abs().idxmin()
    ]
    plt.axvline(x=T_0+1.75, color='black', linestyle='--', label=f'Psid(T_0+1.75)={psi_at_t0_plus_1p75:.3f}')
    plt.fill_betweenx(
        [-psi_peak * 0.2, psi_peak * 0.2],
        T_0 + 1.7,
        T_0 + 1.8,
        color='green',
        alpha=0.3,
        label='20% Psid Peak at T_0+1.75'
    )

    within_20_percent = (abs(psi_at_t0_plus_1p75) <= abs(psi_peak) * 0.2)

    plt.xlabel("Time (s)")
    plt.ylabel("Values")
    plt.title("MF4 Data Reduced Visualization")
    plt.legend()
    plt.grid()

    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    plt.close()

    progress_listbox.insert(tk.END, f"Processed file: {filename}")

    return {
        "filename": filename,
        "plot": buf,
        "psi_peak": psi_peak,
        "within_35_percent": within_35_percent,
        "within_20_percent": within_20_percent,
        "test_passed": within_35_percent and within_20_percent
    }

def create_word_document(data_list, output_filename):
    # Function to create a Word document with analysis results
    doc = Document()
    doc.add_heading('MF4 Data Analysis', 0)

    for data in data_list:
        doc.add_heading(f'File Analysis: {os.path.basename(data["filename"])}', level=1)

        # Add the plot
        doc.add_picture(data["plot"], width=Inches(6))

        # Add psi peak value
        doc.add_paragraph(f'Psi Peak: {data["psi_peak"]:.3f}')

        # Add statements for T_0+1 and T_0+1.75
        doc.add_paragraph(f'Psid at T_0+1 is within 35% range: {data["within_35_percent"]}')
        doc.add_paragraph(f'Psid at T_0+1.75 is within 20% range: {data["within_20_percent"]}')

        # Add test result
        doc.add_paragraph(f'Test Passed: {data["test_passed"]}')

    doc.save(output_filename)
    print("Output saved to", os.path.abspath(output_filename))

def select_folder_and_process():
    # Function to select a folder and process MF4 files
    root = tk.Tk()
    root.title("MF4 File Processor")

    # Create a listbox to show the processing progress
    progress_listbox = Listbox(root, width=80, height=20)
    progress_listbox.pack()

    folder_path = filedialog.askdirectory(title="Select Folder with MF4 Files")
    if not folder_path:
        messagebox.showerror("Error", "No folder selected.")
        root.destroy()
        return

    data_list = []
    for file in os.listdir(folder_path):
        if file.endswith(".mf4"):
            mf4_file = os.path.join(folder_path, file)
            progress_listbox.insert(tk.END, f"Processing file: {mf4_file}")
            root.update_idletasks()  # Update the GUI to show progress

            mf4_data = process_data(mf4_file, convert=True)
            analysis_data = process_mf4_data(mf4_data, mf4_file, progress_listbox)
            if analysis_data:
                data_list.append(analysis_data)

    if data_list:
        output_filename = os.path.join(folder_path, "MF4_Analysis_Report.docx")
        create_word_document(data_list, output_filename)
        progress_listbox.insert(tk.END, f"Analysis complete. Report saved to:\n{os.path.abspath(output_filename)}")
        messagebox.showinfo("Success", f"Analysis complete. Report saved to:\n{os.path.abspath(output_filename)}")
    else:
        progress_listbox.insert(tk.END, "No valid MF4 data processed.")
        messagebox.showinfo("Info", "No valid MF4 data processed.")

    root.mainloop()

if __name__ == "__main__":
    select_folder_and_process()