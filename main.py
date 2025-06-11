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
import pyproj
import utm

# todo: check what happens when we do doc.add... a None

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
            if len(signal.samples) == 0:
                continue
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

def extract_data(filename, convert=False):
    # Function to process the MDF4 file and extract the required signals
    try:
        measurement = MDF(filename)
    except Exception as e:
        print(f"Error opening the file {filename}: {e}")
        return None
    
    signalsInMeas = measurement.channels_db
    
    analysis_channel_list = [
        'Rate_Hor_Z', 
        'AVL_STEA_FTAX_WHL',
        'AVL_STEA_DV',
        'INS_Vel_Hor_X',
        'EgoMobs_Mobs_vx_Act',
        'DcrInEgoM_psid_Act',
        'DcrInEgoM_agwFA_Ste',
        'GNSSPositionDegreeOfLatitude',
        'GNSSPositionDegreeOfLongitude',
        'INS_Lat_Abs_POI1',
        'INS_Long_Abs_POI1',
        'INS_Lat_Abs',
        'INS_Long_Abs',
        'DcrInEgoM_beta_Act',
        'Ins_beta_Act',
        'QU_FN_FDR',
    ]
    
    mappedSignals = {}
    for signal in analysis_channel_list:
        for sub, indexgroup in signalsInMeas.items():
            if sub == signal:
                indexFindings = indexgroup[0]
                group = indexFindings[0]
                index = indexFindings[1]
                mappedSignals[signal] = measurement.get(signal,
                                                        group=group,
                                                        index=index)
    analysis_data_converted = _convert_signals(mappedSignals, convert)
    
    manipulated_signals_channel_list = [
        'EgoWhlRtab_rFL_Whl',
        'EgoWhlRtab_rFR_Whl',
        'EgoWhlRtab_rRL_Whl',
        'EgoWhlRtab_rRR_Whl',
        'EgoWhlRtab_rFL_WhlLut',
        'EgoWhlRtab_rFR_WhlLut',
        'EgoWhlRtab_rRL_WhlLut',
        'EgoWhlRtab_rRR_WhlLut',

        'EgoMobs_Mobs_vx_Act',
        'EgoMobs_Mobs_vy_Act',
        'DcrInEgoM_v_Act',
        'DcrInEgoM_vx_Act',
        'DcrInEgoM_vy_Act',  
    ]

    mappedSignals = {}
    for signal in manipulated_signals_channel_list:
        for sub, indexgroup in signalsInMeas.items():
            if sub == signal:
                indexFindings = indexgroup[0]
                group = indexFindings[0]
                index = indexFindings[1]
                mappedSignals[signal] = measurement.get(signal,
                                                        group=group,
                                                        index=index)
    manipulated_signals_data_converted = _convert_signals(mappedSignals, convert)

    print("Processing analysis data from file:", filename)
    return analysis_data_converted, manipulated_signals_data_converted

def process_mf4_analysis_data(mf4_data, filename, progress_listbox):
    warnings = []
    plot_buf = []

    # Function to process the MF4 data and perform analysis
    if mf4_data is None or mf4_data.empty:
        progress_listbox.insert(tk.END, f"MF4 data is empty or invalid for file: {filename}")
        return None

    # Find signal for psid
    signal_name_psi_d = ""
    unit_psi_d = ""
    if "Rate_Hor_Z" in mf4_data.columns:
        signal_name_psi_d = "Rate_Hor_Z"
        unit_psi_d = "°/s"
    elif "DcrInEgoM_psid_Act" in mf4_data.columns:
        progress_listbox.insert(tk.END, f"Rate_Hor_Z not available for file: {filename}, using DcrInEgoM_psid_Act instead")
        signal_name_psi_d = "DcrInEgoM_psid_Act"
        unit_psi_d = "rad/s"
    else:
        progress_listbox.insert(tk.END, f"WARNING: Rate_Hor_Z and DcrInEgoM_psid_Act not available for file: {filename}")
        warnings.append(f"WARNING: Rate_Hor_Z and DcrInEgoM_psid_Act not available for this file")
        return {
            "filename": filename,
            "plot": None,
            "psid_peak": None,
            "signal_info": None,
            "within_35_percent": None,
            "within_20_percent": None,
            "vx_begin_kmh": None,
            "warning": warnings,
            "test_passed": False
            }

    # Find signal for lenkwinkel
    signal_name_lenkwinkel = ""
    unit_lenkwinkel = ""
    if "DcrInEgoM_agwFA_Ste" in mf4_data.columns:
        signal_name_lenkwinkel = "DcrInEgoM_agwFA_Ste"
        unit_lenkwinkel = "rad"
    elif "AVL_STEA_FTAX_WHL" in mf4_data.columns:
        progress_listbox.insert(tk.END, f"DcrInEgoM_agwFA_Ste not available for file: {filename}, using AVL_STEA_FTAX_WHL instead")
        signal_name_lenkwinkel = "AVL_STEA_FTAX_WHL"
        unit_lenkwinkel = "°"
    elif "AVL_STEA_DV" in mf4_data.columns:
        progress_listbox.insert(tk.END, f"DcrInEgoM_agwFA_Ste and AVL_STEA_FTAX_WHL not available for file: {filename}, using AVL_STEA_DV instead")
        signal_name_lenkwinkel = "AVL_STEA_DV"
        unit_lenkwinkel = "°"
    else:
        progress_listbox.insert(tk.END, f"WARNING: DcrInEgoM_agwFA_Ste, AVL_STEA_FTAX_WHL, and AVL_STEA_DV not available for file: {filename}")
        warnings.append(f"WARNING: DcrInEgoM_agwFA_Ste, AVL_STEA_FTAX_WHL, and AVL_STEA_DV not available for this file")
        return {
            "filename": filename,
            "plot": None,
            "psid_peak": None,
            "signal_info": None,
            "within_35_percent": None,
            "within_20_percent": None,
            "vx_begin_kmh": None,
            "warning": warnings,
            "test_passed": False
            }

    # Find signal for vx
    signal_name_vx = ""
    unit_vx = ""
    if "INS_Vel_Hor_X" in mf4_data.columns:
        signal_name_vx = "INS_Vel_Hor_X"
        unit_vx = "m/s"
    elif "EgoMobs_Mobs_vx_Act" in mf4_data.columns:
        progress_listbox.insert(tk.END, f"INS_Vel_Hor_X not available for file: {filename}, using EgoMobs_Mobs_vx_Act instead")
        signal_name_vx = "EgoMobs_Mobs_vx_Act"
        unit_vx = "m/s"
    else:
        progress_listbox.insert(tk.END, f"WARNING: INS_Vel_Hor_X and EgoMobs_Mobs_vx_Act not available for file: {filename}")
        warnings.append(f"WARNING: INS_Vel_Hor_X and EgoMobs_Mobs_vx_Act not available for file: {filename}")     

    # Find start of SWD (QN_FN_FDR = 512 and gradient of d_lenkwinkel/dt > 0.009)
    start_index = 0
    time_diff = round(mf4_data["time"][1] - mf4_data["time"][0],3) # 2 decs behind 0
    distance_in_index_after_0p1_s = int(0.1 / time_diff)
    for i in range(len(mf4_data[signal_name_lenkwinkel]) - distance_in_index_after_0p1_s): # search +-80 samples when QN_FN_FDR is activated, and where gradient is > 0.009
        if any(mf4_data["QU_FN_FDR"][max(i-80,0):min(i+80,len(mf4_data[signal_name_lenkwinkel]) - distance_in_index_after_0p1_s)]  != 512) and abs(mf4_data[signal_name_lenkwinkel].iloc[i + distance_in_index_after_0p1_s] - mf4_data[signal_name_lenkwinkel].iloc[i]) / 0.1 > 0.44:
            start_index = i
            break
    if start_index == 0:
        progress_listbox.insert(tk.END, f"WARNING: Could not find the start of SWD for file: {filename}")
        warnings.append(f"WARNING: Could not find the start of SWD for this file")
        return {
            "filename": filename,
            "plot": None,
            "psid_peak": None,
            "signal_info": None,
            "within_35_percent": None,
            "within_20_percent": None,
            "vx_begin_kmh": None,
            "warning": warnings,
            "test_passed": False
        }
    mf4_data_reduced = mf4_data.iloc[start_index:].reset_index(drop=True) # Trim data

    # Find maximum psid
    psid_peak_index = mf4_data_reduced[signal_name_psi_d].abs().idxmax()
    psid_peak = mf4_data_reduced[signal_name_psi_d].iloc[psid_peak_index]

    # Find time when sine sweep stopped (compare now value to last value), starting from time where we found psid_max
    T_0 = None
    state = 0
    for i in range(psid_peak_index, len(mf4_data_reduced[signal_name_lenkwinkel]) - distance_in_index_after_0p1_s):
        if state == 0:
            if abs((mf4_data_reduced[signal_name_lenkwinkel].iloc[i] - mf4_data_reduced[signal_name_lenkwinkel].iloc[i+distance_in_index_after_0p1_s])/0.1) <= 0:
                state = 1
        elif state == 1:
            if abs((mf4_data_reduced[signal_name_lenkwinkel].iloc[i] - mf4_data_reduced[signal_name_lenkwinkel].iloc[i+distance_in_index_after_0p1_s])/0.1) > 0.2:
                state = 2
        elif state == 2:
            if abs((mf4_data_reduced[signal_name_lenkwinkel].iloc[i] - mf4_data_reduced[signal_name_lenkwinkel].iloc[i+distance_in_index_after_0p1_s])/0.1) < 0.2:
                T_0 = mf4_data_reduced["time"].iloc[i]
                break

    if T_0 is None:
        progress_listbox.insert(tk.END, f"T_0 could not be determined for file: {filename}")
        warnings.append(f"T_0 could not be determined for this file")
        return {
            "filename": filename,
            "plot": None,
            "psid_peak": None,
            "signal_info": None,
            "within_35_percent": None,
            "within_20_percent": None,
            "vx_begin_kmh": None,
            "warning": warnings,
            "test_passed": False
        }

    tStartSine = mf4_data_reduced["time"].iloc[0]
    T_0 = T_0 - tStartSine

    # Create plots for psid 
    fig_psid, ax_psid = plt.subplots(figsize=(10, 6))
    ax_psid.plot(mf4_data_reduced["time"] - tStartSine, mf4_data_reduced[signal_name_psi_d], label=signal_name_psi_d + f" ({unit_psi_d})")
    ax_psid.axvline(x=T_0, color='red', linestyle='--', label=f'T_0={T_0:.3f}s')
    # plot psid peak
    psid_peak_time = mf4_data_reduced["time"].iloc[psid_peak_index] - tStartSine
    ax_psid.axvline(x=psid_peak_time, color='green', linestyle='--', label=f'Psid Peak (Time={psid_peak_time:.3f}s, Val={psid_peak:.3f} {unit_psi_d})')
    # Criterium 1 for psid
    psid_at_t0_plus_1 = mf4_data_reduced[signal_name_psi_d].iloc[
        (mf4_data_reduced["time"] - tStartSine).sub(T_0 + 1).abs().idxmin()
    ]
    ax_psid.axvline(x=T_0+1, color='blue', linestyle='--', label=f'Psid(T_0+1)={psid_at_t0_plus_1:.3f} {unit_psi_d}')
    ax_psid.fill_betweenx(
        [-psid_peak * 0.35, psid_peak * 0.35],
        T_0 + 0.95,
        T_0 + 1.05,
        color='green',
        alpha=0.3,
        label='35% Psid Peak at T_0+1'
    )
    within_35_percent = (abs(psid_at_t0_plus_1) <= abs(psid_peak) * 0.35)
    # Criteirum 2 for psid
    psid_at_t0_plus_1p75 = mf4_data_reduced[signal_name_psi_d].iloc[
        (mf4_data_reduced["time"] - tStartSine).sub(T_0 + 1.75).abs().idxmin()
    ]
    ax_psid.axvline(x=T_0+1.75, color='black', linestyle='--', label=f'Psid(T_0+1.75)={psid_at_t0_plus_1p75:.3f} {unit_psi_d}')
    ax_psid.fill_betweenx(
        [-psid_peak * 0.2, psid_peak * 0.2],
        T_0 + 1.7,
        T_0 + 1.8,
        color='green',
        alpha=0.6,
        label='20% Psid Peak at T_0+1.75'
    )
    within_20_percent = (abs(psid_at_t0_plus_1p75) <= abs(psid_peak) * 0.2)
    ax_psid.set_ylabel(signal_name_psi_d + f" ({unit_psi_d})") # label and legend for psid plot
    ax_psid.legend()
    ax_psid.grid()
    psid_buf = io.BytesIO() # save to buf
    fig_psid.savefig(psid_buf, format='png')
    plt.close(fig_psid)
    plot_buf.append(psid_buf)

    # Plot lenkwinkel
    fig_lw, ax_lw = plt.subplots(figsize=(10, 6))
    ax_lw.plot(mf4_data_reduced["time"] - tStartSine, mf4_data_reduced[signal_name_lenkwinkel], label=signal_name_lenkwinkel + f" ({unit_lenkwinkel})")
    ax_lw.axvline(x=T_0, color='red', linestyle='--')
    ax_lw.set_ylabel(signal_name_lenkwinkel + f" ({unit_lenkwinkel})")
    ax_lw.set_xlabel("Time (s)")
    ax_lw.legend()
    ax_lw.grid()
    lenkwinkel_buf = io.BytesIO()
    fig_lw.savefig(lenkwinkel_buf, format='png')
    plt.close(fig_lw)
    plot_buf.append(lenkwinkel_buf)

    # Check Querversatz (1.07s after start)
    signal_names_querversatz = [('GNSSPositionDegreeOfLatitude', 'GNSSPositionDegreeOfLongitude'), ('INS_Lat_Abs_POI1', 'INS_Long_Abs_POI1'), ('INS_Lat_Abs', 'INS_Long_Abs')]
    signal_name_querversatz = ""
    unit_querversatz = "m"

    fig_querversatz, ax_querversatz = plt.subplots(figsize=(10, 6))
    querversatz_buf = io.BytesIO()
    
    for signal_name_pair in signal_names_querversatz:
        lat_signal, lon_signal = signal_name_pair
        if lat_signal in mf4_data_reduced.columns and lon_signal in mf4_data_reduced.columns:
        
            # alle in grad
            lat1 = mf4_data_reduced[lat_signal][0]
            lon1 = mf4_data_reduced[lon_signal][0]
            lat2 = mf4_data_reduced[lat_signal]
            lon2 = mf4_data_reduced[lon_signal]

            if any(abs(lat2) > 90) or any(abs(lon2) > 180):
                warnings.append(f"WARNING: {lat_signal},{lon_signal} have unvalid values")
                continue

            signal_name_querversatz += f"{lat_signal},{lon_signal},"
            
            ########################## IN HOUSE HAVERSINE ########################
            # lat1 = np.deg2rad(lat1)
            # lon1 = np.deg2rad(lon1)
            # lat2 = np.deg2rad(lat2)
            # lon2 = np.deg2rad(lon2)

            # a = np.sin((lat2 - lat1) / 2) ** 2 + np.cos(lat1) * np.cos(lat2) * np.sin((lon2 - lon1) / 2) ** 2
            # dist = 6378.388 * 2.0 * np.arctan2(np.sqrt(a), np.sqrt(1.0-a)) * 1000 # in m
            
            # num = np.sin(lon2 - lon1) * np.cos(lat2)
            # den = np.cos(lat1) * np.sin(lat2) - np.sin(lat1) * np.cos(lat2) * np.cos(lon2 - lon1)
            # bearing = np.arctan2(num, den)

            # x = dist * np.sin(bearing)
            # y = dist * np.cos(bearing)
            ########################## IN HOUSE HAVERSINE ########################

            ########################## pyproj ####################################
            wgs84 = pyproj.CRS('EPSG:4326')
            # Automatically determine UTM zone using the first coordinate
            utm_zone = utm.from_latlon(lat1, lon1)
            utm_crs = pyproj.CRS(f'EPSG:{32600 + utm_zone[2]}')

            # Create a transformer from WGS84 to the determined UTM zone
            transformer = pyproj.Transformer.from_crs(wgs84, utm_crs)

            # Convert the array of positions to UTM
            easthing, northing = transformer.transform(lat2, lon2)
            y = easthing - easthing[0]
            x = northing - northing[0]
            ########################## pyproj ####################################

            initial_heading = np.arctan2(y[1] - y[0], x[1] - x[0])
            lateral = (x - x[0]) * np.sin(initial_heading) - (y - y[0]) * np.cos(initial_heading)
            longitudinal = (x - x[0]) * np.cos(initial_heading) + (y - y[0]) * np.sin(initial_heading)
            index_lenkwinkel_1p07_after = np.argmin(np.abs(mf4_data_reduced["time"] - (tStartSine + 1.07)))
            lateral_1p07 = lateral[index_lenkwinkel_1p07_after]

            ax_querversatz.plot(mf4_data_reduced["time"] - tStartSine, lateral)
            ax_querversatz.scatter(mf4_data_reduced["time"].iloc[index_lenkwinkel_1p07_after] - tStartSine, lateral_1p07, marker='x', color='red', label=f"Lateral Displacement at 1.07s: {lateral_1p07:.2f} {unit_querversatz}")
            ax_querversatz.set_xlabel("Time (s)")
            ax_querversatz.set_ylabel("Querversatz from " + lat_signal + "," + lon_signal + f" ({unit_querversatz})")
            ax_querversatz.grid()
            ax_querversatz.legend()
            fig_querversatz.savefig(querversatz_buf, format='png')
            plt.close(fig_querversatz)
            

            if lateral_1p07 < 1.83:
                warnings.append(f"WARNING: Querversatz 1.07s after start of SWD is smaller than 1,83 meter for signal pair {signal_name_pair}")
                break
            
    if not signal_name_querversatz:
        plt.close(fig_querversatz)
        progress_listbox.insert(tk.END, f"WARNING: Querversatz not available for file: {filename}")
        warnings.append("WARNING: Querversatz not available for this file")
    else:
        plot_buf.append(querversatz_buf)
    
    # Check vx in km/h
    vx_begin_kmh = mf4_data_reduced[signal_name_vx][0] * 3.6
    vx_within_80_pm_2_kmh = (78 <= vx_begin_kmh <= 82)

    # Check Schwimmwinkel
    signal_names_schwimmwinkel = ["DcrInEgoM_beta_Act", "Ins_beta_Act"]
    units_schwimmwinkel = ["rad", "rad"]
    signal_name_schwimmwinkel = ""
    unit_schwimmwinkel = ""

    for signal_name, unit in zip(signal_names_schwimmwinkel, units_schwimmwinkel):
        if signal_name in mf4_data_reduced.columns:
            signal_name_schwimmwinkel += signal_name + ","
            unit_schwimmwinkel += unit + ","
            schwimmwinkel = mf4_data_reduced[signal_name]
            if unit == "rad":
                schwimmwinkel = np.rad2deg(schwimmwinkel)
            if np.any(np.abs(schwimmwinkel) > 15):
                warnings.append(f"WARNING: Schwimmwinkel is bigger than 15 degrees for signal {signal_name}, maximum absolute value: {np.max(abs(schwimmwinkel))} degrees")
    if not signal_name_schwimmwinkel:
        progress_listbox.insert(tk.END, f"WARNING: Schwimmwinkel not available for file: {filename}")
        warnings.append("WARNING: Schwimmwinkel not available for this file")

    # Criterium for test passed
    test_passed = within_35_percent and within_20_percent and vx_within_80_pm_2_kmh

    # Create a dictionary to store signals used
    signal_info = {
        "psid": {"name": signal_name_psi_d, "unit": unit_psi_d},
        "lenkwinkel": {"name": signal_name_lenkwinkel, "unit": unit_lenkwinkel},
        "vx": {"name": signal_name_vx, "unit": unit_vx},  # still included for completeness
        "schwimmwinkel": {"name": signal_name_schwimmwinkel, "unit": unit_schwimmwinkel},
        "querversatz": {"name": signal_name_querversatz, "unit": unit_querversatz}
    }
    
    progress_listbox.insert(tk.END, f"Processed file: {filename}")

    return {
        "filename": filename,
        "plot": plot_buf,
        "psid_peak": psid_peak,
        "signal_info": signal_info,
        "within_35_percent": within_35_percent,
        "within_20_percent": within_20_percent,
        "vx_begin_kmh": vx_begin_kmh,
        "warning": warnings,
        "test_passed": test_passed
    }

def process_mf4_manipulated_signals_data(mf4_data, filename, progress_listbox):
    plot_buf = []

    # Function to process the manipulated signals data
    if mf4_data is None or mf4_data.empty:
        progress_listbox.insert(tk.END, f"Manipulated signals data is empty or invalid for file: {filename}")
        return {
            "filename": filename,
            "plot": None,
        }

    # Check whether rDyn is manipulated -------------------------------------------------------------------
    fig_rDyn, ax_rDyn = plt.subplots(figsize=(10, 6))
    rDyn_is_manipulated = False
    if "EgoWhlRtab_rFL_Whl" in mf4_data.columns and "EgoWhlRtab_rFL_WhlLut" in mf4_data.columns:
        if any(abs(mf4_data["EgoWhlRtab_rFL_Whl"] - mf4_data["EgoWhlRtab_rFL_WhlLut"]) >= 0.03):
            ax_rDyn.plot(mf4_data['time'], mf4_data['EgoWhlRtab_rFL_Whl'], label='EgoWhlRtab_rFL_Whl')
            ax_rDyn.plot(mf4_data['time'], mf4_data['EgoWhlRtab_rFL_WhlLut'], label='EgoWhlRtab_rFL_WhlLut')
            rDyn_is_manipulated = True
    if "EgoWhlRtab_rFR_Whl" in mf4_data.columns and "EgoWhlRtab_rFR_WhlLut" in mf4_data.columns:
        if any(abs(mf4_data["EgoWhlRtab_rFR_Whl"] - mf4_data["EgoWhlRtab_rFR_WhlLut"]) >= 0.03):
            ax_rDyn.plot(mf4_data['time'], mf4_data['EgoWhlRtab_rFR_Whl'], label='EgoWhlRtab_rFR_Whl')
            ax_rDyn.plot(mf4_data['time'], mf4_data['EgoWhlRtab_rFR_WhlLut'], label='EgoWhlRtab_rFLR_WhlLut')
            rDyn_is_manipulated = True
    if "EgoWhlRtab_rRL_Whl" in mf4_data.columns and "EgoWhlRtab_rRL_WhlLut" in mf4_data.columns:
        if any(abs(mf4_data["EgoWhlRtab_rRL_Whl"] - mf4_data["EgoWhlRtab_rRL_WhlLut"]) >= 0.03):
            ax_rDyn.plot(mf4_data['time'], mf4_data['EgoWhlRtab_rRL_Whl'], label='EgoWhlRtab_rRL_Whl')
            ax_rDyn.plot(mf4_data['time'], mf4_data['EgoWhlRtab_rRL_WhlLut'], label='EgoWhlRtab_rRL_WhlLut')
            rDyn_is_manipulated = True
    if "EgoWhlRtab_rRR_Whl" in mf4_data.columns and "EgoWhlRtab_rRR_WhlLut" in mf4_data.columns:
        if any(abs(mf4_data["EgoWhlRtab_rRR_Whl"] - mf4_data["EgoWhlRtab_rRR_WhlLut"]) >= 0.03):
            ax_rDyn.plot(mf4_data['time'], mf4_data['EgoWhlRtab_rRR_Whl'], label='EgoWhlRtab_rRR_Whl')
            ax_rDyn.plot(mf4_data['time'], mf4_data['EgoWhlRtab_rRR_WhlLut'], label='EgoWhlRtab_rRR_WhlLut')
            rDyn_is_manipulated = True
    if rDyn_is_manipulated:
        ax_rDyn.set_xlabel("Time (s)")
        ax_rDyn.set_ylabel("rDyn (m)")
        ax_rDyn.legend()
        ax_rDyn.grid()
        rDyn_buf = io.BytesIO()
        fig_rDyn.savefig(rDyn_buf, format='png')
        plt.close(fig_rDyn)
        plot_buf.append(rDyn_buf)
    
    # Check whether rDyn is manipulated ----------------------------------------------------------------------
    
    # Check whether vx or vy is manipulated ------------------------------------------------------------------
        # 'EgoMobs_Mobs_vx_Act',
        # 'EgoMobs_Mobs_vy_Act',
        # 'DcrInEgoM_v_Act',
        # 'DcrInEgoM_vx_Act',
        # 'DcrInEgoM_vy_Act',  
    fig_v, ax_v = plt.subplots(figsize=(10, 6))
    v_is_manipulated = False
    if "EgoMobs_Mobs_vx_Act" in mf4_data.columns and "DcrInEgoM_vx_Act" in mf4_data.columns:
        difference = np.abs(mf4_data["EgoMobs_Mobs_vx_Act"] - mf4_data["DcrInEgoM_vx_Act"])
        index = np.where(np.diff(difference) >= 0.2)[0]
        first_index = index[0] if index.size > 0 else None
        if first_index is not None:
            ax_v.plot(mf4_data['time'][first_index - 10:first_index + 10], mf4_data['EgoMobs_Mobs_vx_Act'][first_index - 10:first_index + 10], label='EgoMobs_Mobs_vx_Act')
            ax_v.plot(mf4_data['time'][first_index - 10:first_index + 10], mf4_data['DcrInEgoM_vx_Act'][first_index - 10:first_index + 10], label='DcrInEgoM_vx_Act')
            v_is_manipulated = True
    if "EgoMobs_Mobs_vy_Act" in mf4_data.columns and "DcrInEgoM_vy_Act" in mf4_data.columns:
        difference = np.abs(mf4_data["EgoMobs_Mobs_vy_Act"] - mf4_data["DcrInEgoM_vy_Act"])
        index = np.where(np.diff(difference) >= 0.2)[0]
        first_index = index[0] if index.size > 0 else None
        if first_index is not None:
            ax_v.plot(mf4_data['time'][first_index - 10:first_index + 10], mf4_data['EgoMobs_Mobs_vy_Act'][first_index - 10:first_index + 10], label='EgoMobs_Mobs_vx_Act')
            ax_v.plot(mf4_data['time'][first_index - 10:first_index + 10], mf4_data['DcrInEgoM_vy_Act'][first_index - 10:first_index + 10], label='DcrInEgoM_vx_Act')
            v_is_manipulated = True
    if v_is_manipulated:
        ax_v.set_xlabel("Time (s)")
        ax_v.set_ylabel("Velocity (m/s)")
        ax_v.legend()
        ax_v.grid()
        v_buf = io.BytesIO()
        fig_v.savefig(v_buf, format='png')
        plt.close(fig_rDyn)
        plot_buf.append(v_buf)
    # Check whether vx or vy is manipulated ------------------------------------------------------------------
    return {
        "filename": filename,
        "plot": plot_buf,
    }


def create_word_document(analysis_data_list, manipulated_signals_data_list, output_filename):
    # Function to create a Word document with analysis results
    doc = Document()
    doc.add_heading('MF4 Data Analysis', 0)

    for (analysis_data, manipulated_signals_data) in zip(analysis_data_list, manipulated_signals_data_list):
        doc.add_heading(f'File Analysis: {os.path.basename(analysis_data["filename"])}', level=1)

        if analysis_data["warning"] is not None:
            for warning in analysis_data["warning"]:
                doc.add_paragraph(warning, style='ListBullet')

        if analysis_data["plot"] is not None:
            for plot_buf in analysis_data["plot"]:
                plot_buf.seek(0)  # Reset the file pointer to the beginning
                doc.add_picture(plot_buf, width=Inches(6))
                doc.add_paragraph()
        
        if manipulated_signals_data["plot"] is not None:
            doc.add_paragraph('Manipulated Signals:')
            for plot_buf in manipulated_signals_data["plot"]:
                plot_buf.seek(0)
                doc.add_picture(plot_buf, width=Inches(6))
                doc.add_paragraph()

        if analysis_data["signal_info"] is not None:
            doc.add_paragraph(f'Psid Signal Used: {analysis_data["signal_info"]["psid"]["name"]} ({analysis_data["signal_info"]["psid"]["unit"]})')

            doc.add_paragraph(f'Lenkwinkel Signal Used: {analysis_data["signal_info"]["lenkwinkel"]["name"]} ({analysis_data["signal_info"]["lenkwinkel"]["unit"]})')

            doc.add_paragraph(f'Vx Signal Used: {analysis_data["signal_info"]["vx"]["name"]} ({analysis_data["signal_info"]["vx"]["unit"]})')

            doc.add_paragraph(f'Schwimmwinkel Signal Used: {analysis_data["signal_info"]["schwimmwinkel"]["name"]} ({analysis_data["signal_info"]["schwimmwinkel"]["unit"]})')

            doc.add_paragraph(f'Querversatz Signal Used: {analysis_data["signal_info"]["querversatz"]["name"]} ({analysis_data["signal_info"]["querversatz"]["unit"]})')

            doc.add_paragraph(f'Psid Peak: {analysis_data["psid_peak"]:.3f} {analysis_data["signal_info"]["psid"]["unit"]}')

        if analysis_data["within_35_percent"] is not None:
            doc.add_paragraph(f'Psid at T_0+1 is within 35% range: {analysis_data["within_35_percent"]}')

        if analysis_data["within_20_percent"] is not None:
            doc.add_paragraph(f'Psid at T_0+1.75 is within 20% range: {analysis_data["within_20_percent"]}')

        if analysis_data["vx_begin_kmh"] is not None:
            doc.add_paragraph(f'Vx at in the beginning of SWD: {analysis_data["vx_begin_kmh"]:.3f} km/h')

        if analysis_data["test_passed"] is not None:
            doc.add_paragraph(f'Test Passed: {analysis_data["test_passed"]}')

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

    analysis_data_list = []
    manipulated_signals_data_list = []
    for file in os.listdir(folder_path):
        if file.endswith(".mf4"):
            mf4_file = os.path.join(folder_path, file)
            progress_listbox.insert(tk.END, f"Processing file: {mf4_file}")
            root.update_idletasks()  # Update the GUI to show progress

            mf4_analysis_data, mf4_manipulated_signals_data = extract_data(mf4_file, convert=True)
            analysis_data = process_mf4_analysis_data(mf4_analysis_data, mf4_file, progress_listbox)
            manipulated_signals_data = process_mf4_manipulated_signals_data(mf4_manipulated_signals_data, mf4_file, progress_listbox)
            if analysis_data is not None:
                analysis_data_list.append(analysis_data)
            if manipulated_signals_data is not None:
                manipulated_signals_data_list.append(manipulated_signals_data)

    if len(analysis_data_list) > 0 and len(manipulated_signals_data_list) > 0:
        output_filename = os.path.join(folder_path, "MF4_Analysis_Report.docx")
        create_word_document(analysis_data_list, manipulated_signals_data_list,output_filename)
        progress_listbox.insert(tk.END, f"Analysis complete. Report saved to:\n{os.path.abspath(output_filename)}")
        messagebox.showinfo("Success", f"Analysis complete. Report saved to:\n{os.path.abspath(output_filename)}")
    else:
        progress_listbox.insert(tk.END, "No valid MF4 data processed.")
        messagebox.showinfo("Info", "No valid MF4 data processed.")

    root.mainloop()

if __name__ == "__main__":
    select_folder_and_process()