import pandas as pd
from datetime import timedelta
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Plate Simulation App", layout="centered")
st.title("📊 Plate Processing Simulation")

uploaded_file = st.file_uploader("Upload Excel File (Sheet1 must be present)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="Sheet1")
    df.columns = df.columns.str.strip()
    df['Rolling Time'] = pd.to_datetime(df['Rolling Time'])
    df['Index'] = df.index

    columns = [
        'Start Cooling', 'End Cooling', 'Start Shearing', 'End Shearing',
        'Trimming Station', 'Start Trimming', 'End Trimming', 'UT Station',
        'Start UT', 'End UT', 'Start Punching', 'End Punching', 'Start Inspection',
        'End Inspection', 'Norm Ready Time', 'Furnace', 'Start Normalizing',
        'End Normalizing', 'Start Levelling', 'End Levelling',
        'Start Final Inspection', 'End Final Inspection', 'Finish Time'
    ]
    for col in columns:
        df[col] = None

    priority_groups = {
        '<=40': ['Oxy1', 'Plasma1', 'Pug1', 'Pug2'],
        '>40': ['Oxy2', 'Pug3', 'Pug4', 'Pug5', 'Pug6', 'Pug7']
    }
    cutting_durations = {m: 35 if 'Oxy' in m or 'Plasma' in m else 90 for group in priority_groups.values() for m in group}
    machine_avail = {m: df['Rolling Time'].min() for m in cutting_durations}
    ut_machines = [df['Rolling Time'].min()] * 3
    final_buffer = timedelta(hours=4)
    punch_time = timedelta(minutes=5)

    def cooling_time(thk): return timedelta(hours=8 if thk < 40 else 24)

    norm_inputs = []

    for i, row in df.iterrows():
        df.at[i, 'Start Cooling'] = row['Rolling Time']
        df.at[i, 'End Cooling'] = row['Rolling Time'] + cooling_time(row['Thickness'])
        time = df.at[i, 'End Cooling']

        if row['Thickness'] >= 40:
            df.at[i, 'Start Shearing'] = time
            df.at[i, 'End Shearing'] = time + timedelta(hours=1)
            time = df.at[i, 'End Shearing']

        if 'Trimmed' in str(row['Edge Condition']):
            group = '<=40' if row['Thickness'] <= 40 else '>40'
            machines = priority_groups[group]
            best = min(machines, key=lambda m: machine_avail[m])
            start = max(time, machine_avail[best])
            duration = timedelta(minutes=cutting_durations[best])
            end = start + duration
            df.at[i, 'Trimming Station'] = best
            df.at[i, 'Start Trimming'] = start
            df.at[i, 'End Trimming'] = end
            machine_avail[best] = end
            time = end

        if pd.notna(row['UT']) and str(row['UT']).strip():
            idx = ut_machines.index(min(ut_machines))
            start_ut = max(time, ut_machines[idx])
            df.at[i, 'UT Station'] = f'UT{idx+1}'
            df.at[i, 'Start UT'] = start_ut
            df.at[i, 'End UT'] = start_ut + timedelta(minutes=15)
            ut_machines[idx] = df.at[i, 'End UT']
            time = df.at[i, 'End UT']

        df.at[i, 'Start Punching'] = time
        df.at[i, 'End Punching'] = time + punch_time
        time = df.at[i, 'End Punching']

        df.at[i, 'Start Inspection'] = time
        df.at[i, 'End Inspection'] = time + timedelta(minutes=5)
        time = df.at[i, 'End Inspection']

        if str(row['Supply Condition']).strip().lower() == 'normalized' and row['Thickness'] >= 14:
            norm_ready = time + timedelta(hours=3)
            df.at[i, 'Norm Ready Time'] = norm_ready
            norm_inputs.append({'Index': i, 'Base Time': norm_ready, 'Thickness': row['Thickness']})
        else:
            df.at[i, 'Finish Time'] = time + final_buffer

    def normalize_discrete_height(norm_df):
        df_norm = pd.DataFrame(norm_df)
        df_norm['Grouped'] = False
        df_norm['Furnace'] = None
        df_norm['Norm Start'] = None
        df_norm['Norm End'] = None
        df_norm['Level Start'] = None
        df_norm['Level End'] = None

        avail = {'NF1': df['Rolling Time'].min(), 'NF2': df['Rolling Time'].min()}
        capacity = {'NF1': 1630, 'NF2': 800}
        cycle = timedelta(hours=12)
        cooldown = timedelta(hours=6)
        next_furnace = 'NF1'

        while not df_norm[df_norm['Grouped'] == False].empty:
            f = next_furnace
            cap = capacity[f]
            ready = avail[f]
            stack, total, latest_ready = [], 0, ready

            for idx, row in df_norm[df_norm['Grouped'] == False].sort_values(by='Base Time').iterrows():
                h = row['Thickness']
                if h and total + h <= cap:
                    stack.append(idx)
                    total += h
                    latest_ready = max(latest_ready, row['Base Time'])

            if stack:
                start = latest_ready
                end = start + cycle
                df_norm.loc[stack, ['Grouped', 'Furnace', 'Norm Start', 'Norm End']] = True, f, start, end
                avail[f] = end + cooldown
                next_furnace = 'NF2' if f == 'NF1' else 'NF1'
            else:
                break

        df_norm['Norm End'] = pd.to_datetime(df_norm['Norm End'], errors='coerce')
        leveller_time = df_norm['Norm End'].dropna().min() + timedelta(hours=8)

        for idx, row in df_norm.iterrows():
            if pd.isna(row['Norm End']):
                continue
            start = max(row['Norm End'] + timedelta(hours=8), leveller_time)
            end = start + timedelta(minutes=10)
            df_norm.at[idx, 'Level Start'] = start
            df_norm.at[idx, 'Level End'] = end
            leveller_time = end

        return df_norm

    norm_inputs_df = pd.DataFrame([n for n in norm_inputs if str(df.at[n['Index'], 'Supply Condition']).strip().lower() == 'normalized'])
    norm_inputs_df['Index'] = norm_inputs_df['Index'].astype(int)
    norm_result = normalize_discrete_height(norm_inputs_df)

    for _, r in norm_result.iterrows():
        i = r['Index']
        thk = int(df.at[i, 'Thickness'])
        df.at[i, 'Furnace'] = r['Furnace']
        df.at[i, 'Start Normalizing'] = r['Norm Start']
        df.at[i, 'End Normalizing'] = r['Norm End']
        df.at[i, 'Start Levelling'] = r['Level Start']
        df.at[i, 'End Levelling'] = r['Level End']
        if pd.notna(r['Level End']):
            df.at[i, 'Start Final Inspection'] = r['Level End']
            df.at[i, 'End Final Inspection'] = r['Level End'] + timedelta(minutes=thk * 2)
            df.at[i, 'Finish Time'] = df.at[i, 'End Final Inspection'] + final_buffer

    df.drop(columns=['Index'], inplace=True)
    output = BytesIO()
    df.to_excel(output, index=False)
    st.success("✅ Simulation completed.")
    st.download_button("📥 Download Result", output.getvalue(), file_name="Final_Simulation_Corrected_Clean.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
