import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from asammdf import MDF
from docx import Document
from docx.shared import Inches
import io

def _get_time_range(signals: dict) -> np.ndarray:
    # Funktion um den Start- und den Endpunkt der Messung zu bestimmen
    # beliebiger Wert für timeStart
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
    # print(f"   Startzeit: {timeStart:.4f}, Endzeit: {timeEnd:.4f}")
    return timeVector


def _convert_signals(signals: dict, convert:bool):
    # die funktion convertiert asammdf signale
    # auf einheitlichen Zeitvektor
    time_vector = _get_time_range(signals)
    if convert:#return Datafram
        df = pd.DataFrame(time_vector, columns=['time'])
        for key, signal in signals.items():
            syncSignal = signal.interp(time_vector, 0, 0)
            dataSignal = list(zip(syncSignal.timestamps, syncSignal.samples))
            dfSignal = pd.DataFrame(dataSignal, columns=['time', key])
            df = pd.merge(df, dfSignal, on='time')
        return df
    else:#return dict
        for key, signal in signals.items():                
            syncSignal = signal.interp(time_vector, 0, 0)
            signal.samples = syncSignal.samples
        signals['time'] = time_vector
        return signals

def process_data(filename, convert=False):
    try:
        # Öffnen der MDF4-Datei
        measurement = MDF(filename)
    except Exception as e:
        print(f"Fehler beim Öffnen der Datei {filename}: {e}")
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
    convertedSignals = _convert_signals(mappedSignals,convert)
    print("Processing data from file:", filename)
    return convertedSignals

def process_mf4_data(mf4_data, filename):
    if mf4_data is None or mf4_data.empty:
        print("MF4 data is empty or invalid.")
        return None

    # State machine to find the first non-512 value from the end, then where it is 512 again
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
        print("Could not find the required transition in QU_FN_FDR.")
        return None

    # Search 30 indexes before index_qu_fn_fdr where the difference in DcrInEgoM_psid_Act exceeds 0.008
    index_dcr_in_ego = 2
    for i in range(index_qu_fn_fdr - 30, index_qu_fn_fdr):
        if i > 0 and abs(mf4_data["DcrInEgoM_agwFA_Ste"].iloc[i] - mf4_data["DcrInEgoM_agwFA_Ste"].iloc[i - 1]) > 0.009:
            index_dcr_in_ego = i - 2
            break

    # Set the first time index for mf4
    start_index = min(index_dcr_in_ego, index_dcr_in_ego)
    mf4_data_reduced = mf4_data.iloc[start_index:].reset_index(drop=True)

    # Find the signed peak value in DcrInEgoM_psid_Act
    psi_peak = mf4_data_reduced["DcrInEgoM_psid_Act"].iloc[
        mf4_data_reduced["DcrInEgoM_psid_Act"].abs().idxmax()
    ]

    # Find T_0 by searching from the end where abs(diff(DcrInEgoM_agwFA_Ste)) > 0.002
    T_0 = None
    for i in range(len(mf4_data_reduced["DcrInEgoM_agwFA_Ste"]) - 1, 0, -1):
        if abs(mf4_data_reduced["DcrInEgoM_agwFA_Ste"].iloc[i] - mf4_data_reduced["DcrInEgoM_agwFA_Ste"].iloc[i - 1]) > 0.002:
            T_0 = mf4_data_reduced["time"].iloc[i]
            break
    if T_0 is None:
        print(f"T_0 nto found for file {filename}")
        return None

    # Plot the data starting from sine wave
    tStartSine = mf4_data_reduced["time"].iloc[0]
    time = mf4_data_reduced["time"] - tStartSine  # Normalize time to start from 0
    T_0 = T_0 - tStartSine  # Normalize T to start from 0
    plt.figure(figsize=(10, 6))
    for column in mf4_data_reduced.columns:
        if column != "time" and column != "QU_FN_FDR":
            plt.plot(time, mf4_data_reduced[column], label=column)
    
    # Add markers for T_0 and theta_peak
    if T_0 is not None:
        plt.axvline(x=T_0, color='red', linestyle='--', label=f'T_0={T_0:.3f}s')
    psi_peak_time = mf4_data_reduced["time"].iloc[
        mf4_data_reduced["DcrInEgoM_psid_Act"].abs().idxmax()
    ] - tStartSine
    plt.axvline(x=psi_peak_time, color='green', linestyle='--', label=f'Psid Peak (Time={psi_peak_time:.3f}s, Val={psi_peak:.3f})')

    # Find the value of psi at T_0+1
    psi_at_t0_plus_1 = mf4_data_reduced["DcrInEgoM_psid_Act"].iloc[
        (mf4_data_reduced["time"] - tStartSine).sub(T_0 + 1).abs().idxmin()
    ]
    # Add semi-transparent green box at T+1, where height is 35% of psi_peak
    plt.axvline(x=T_0+1, color='black', linestyle='--', label=f'Psid(T_0+1)={psi_at_t0_plus_1:.3f}')
    plt.fill_betweenx(
        [-psi_peak * 0.35, psi_peak * 0.35],  # Y-range for the box
        T_0 + 0.95,  # Start of the box
        T_0 + 1.05,  # End of the box
        color='green',
        alpha=0.3,
        label='35% Psid Peak at T_0+1'
    )

    # Check if the value at T_0+1 is within 35% of psi_peak
    within_35_percent = (abs(psi_at_t0_plus_1) <= abs(psi_peak) * 0.35)

    # Find the value of psi at T_0+1.75
    psi_at_t0_plus_1p75 = mf4_data_reduced["DcrInEgoM_psid_Act"].iloc[
        (mf4_data_reduced["time"] - tStartSine).sub(T_0 + 1.75).abs().idxmin()
    ]
    # Add semi-transparent green box at T+1.75, where height is 20% of psi_peak
    plt.axvline(x=T_0+1.75, color='black', linestyle='--', label=f'Psid(T_0+1.75)={psi_at_t0_plus_1p75:.3f}')
    plt.fill_betweenx(
        [-psi_peak * 0.2, psi_peak * 0.2],  # Y-range for the box
        T_0 + 1.7,  # Start of the box
        T_0 + 1.8,  # End of the box
        color='green',
        alpha=0.3,
        label='20% Psid Peak at T_0+1.75'
    )

    # Check if the value at T_0+1.75 is within 20% of psi_peak
    within_20_percent = (abs(psi_at_t0_plus_1p75) <= abs(psi_peak) * 0.2)

    plt.xlabel("Time (s)")
    plt.ylabel("Values")
    plt.title("MF4 Data Reduced Visualization")
    plt.legend()
    plt.grid()

    # Save the plot to a bytes object
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    plt.close()

    return {
        "filename": filename,
        "plot": buf,
        "psi_peak": psi_peak,
        "within_35_percent": within_35_percent,
        "within_20_percent": within_20_percent,
        "test_passed": within_35_percent and within_20_percent
    }

def create_word_document(data_list, output_filename):
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

    print("Output saved to", output_filename)

if __name__ == "__main__":
    # Initialize an empty list to store dataframes
    data_list = []
    # Path to the folder containing the files
    folder_path = "Messungen_2025-04-22_V141959_VS0_NA5_LR_AWD_20Z_Winter_SWD_185_Links_SZ8_Manip"
    # Iterate through the folder and find matching .dcm and .mf4 files
    for file in os.listdir(folder_path):
        if file.endswith(".mf4"):
            base_name = os.path.splitext(file)[0]
            mf4_file = os.path.join(folder_path, f"{base_name}.mf4")
            dcm_file = os.path.join(folder_path, f"{base_name}_IPF_FAR.dcm")

            if os.path.exists(dcm_file):
                # Parse the .dcm file
                # dcm_dict = parse_dcm_file(dcm_file)
                dcm_dict = None

                # Parse the .mf4 file
                mf4_data = process_data(mf4_file, convert=True)

                # Process the mf4 data
                analysis_data = process_mf4_data(mf4_data, mf4_file)
                if analysis_data:
                    data_list.append(analysis_data)

    # Create the Word document with the analysis data
    create_word_document(data_list, "MF4_Analysis_Report.docx")