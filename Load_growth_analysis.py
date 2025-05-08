import socket
import time
import matplotlib.pyplot as plt
import pandas as pd

def connect(port_number):
    TCP_IP = '127.0.0.1'
    TCP_PORT = port_number
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.connect((TCP_IP, TCP_PORT))
    return s

def change_slider_val(s, slider_name, slider_val):
    message = f'SetSlider "Subsystem #1 : CTLs : Inputs : {slider_name}" = {slider_val};'
    s.send(message.encode())

def get_meter_val(s, meter_name):
    message = f'float temp_float = MeterCapture("{meter_name}");'
    BUFFER_SIZE = 2048
    s.send(message.encode())
    s.send('string temp_string;'.encode())
    s.send('sprintf(temp_string, "%f", temp_float);'.encode())
    s.send('ListenOnPortHandshake(temp_string);'.encode())
    tokenstring = s.recv(BUFFER_SIZE)
    return float(tokenstring[:len(tokenstring)-2])

loads = {
    "LoadI1": {"slider": "SLI1", "P": "PI1", "Q": "QI1"},
    "LoadI2": {"slider": "SLI2", "P": "PI2", "Q": "QI2"},
    "LoadC2": {"slider": "SLC2", "P": "PC2", "Q": "QC2"},
    "LoadP1": {"slider": "SLP1", "P": "PP1", "Q": "QP1"}
}

bus_voltage_meters = {
    "Bus101A": "V101A_rms",
    "Bus102A": "V102A_rms",
    "Bus105A": "V105A_rms", 
    "Bus106A": "V106A_rms", 
    "Bus107A": "V107A_rms",
 
}

port = 4505
base_load = 1.0
growth_rate = 0.05
years = 5
wait_time = 1

s = connect(port)
all_data = []

for year in range(1, years + 1):
    year_row = {"Year": year}
    total_load_p = 0

    for load_name, config in loads.items():
        load_val = base_load * (1 + growth_rate) ** (year - 1)
        change_slider_val(s, config["slider"], load_val)
        time.sleep(wait_time)

        p = get_meter_val(s, config["P"])
        q = get_meter_val(s, config["Q"])
        total_load_p += p

        year_row[f"{load_name}_Load"] = load_val
        year_row[f"{load_name}_P"] = p
        year_row[f"{load_name}_Q"] = q

        print(f"[Year {year}] {load_name} → Load: {load_val:.2f}, P: {p:.3f}, Q: {q:.3f}")

    # Read bus voltages
    for label, meter in bus_voltage_meters.items():
        voltage = get_meter_val(s, meter)
        year_row[label] = voltage
        print(f"[Year {year}] {label}: {voltage:.3f}")

    # Read P_Grid 
    p_grid = get_meter_val(s, "P_Grid")
    year_row["P_Grid"] = p_grid
    year_row["Total_Load_P"] = total_load_p
    year_row["Grid_Efficiency"] = total_load_p / p_grid if p_grid != 0 else 0
    print(f"[Year {year}] P_Grid: {p_grid:.3f} MW, Efficiency: {year_row['Grid_Efficiency']:.3f}")

    all_data.append(year_row)

#Save results to Excel 
df = pd.DataFrame(all_data)
save_path = "C:/Users/CESAC/Desktop/Bijay/Data/load_growth_results.xlsx"
df.to_excel(save_path, index=False)
print(f"\n Excel file saved at: {save_path}")

# Plot all metrics
def plot_all_on_one_graph(df):
    fig, axs = plt.subplots(3, 2, figsize=(16, 12), sharex=True)
    axs = axs.flatten()  # Flatten for easy indexing

    # Real Power
    p_cols = [col for col in df.columns if col.endswith('_P') and not col.startswith("Total")]
    for col in p_cols:
        axs[0].plot(df["Year"], df[col], marker='o', label=col)
    axs[0].set_title("Real Power (P) vs Year")
    axs[0].set_ylabel("Power (MW)")
    axs[0].legend()
    axs[0].grid(True)

    # Reactive Power
    q_cols = [col for col in df.columns if col.endswith('_Q')]
    for col in q_cols:
        axs[1].plot(df["Year"], df[col], marker='o', label=col)
    axs[1].set_title("Reactive Power (Q) vs Year")
    axs[1].set_ylabel("Reactive Power (MVAr)")
    axs[1].legend()
    axs[1].grid(True)

    # Bus Voltages
    bus_cols = [col for col in df.columns if col.startswith("Bus")]
    for col in bus_cols:
        axs[2].plot(df["Year"], df[col], marker='o', label=col)
    axs[2].set_title("Bus Voltages vs Year")
    axs[2].set_ylabel("Voltage (p.u.)")
    axs[2].legend()
    axs[2].grid(True)

    # Grid Power
    axs[3].plot(df["Year"], df["P_Grid"], marker='o', color='purple', label="P_Grid (MW)")
    axs[3].set_title("Grid Power vs Year")
    axs[3].set_ylabel("P_Grid (MW)")
    axs[3].legend()
    axs[3].grid(True)

    # Grid Efficiency
    axs[4].plot(df["Year"], df["Grid_Efficiency"], marker='o', color='green', label="Grid Efficiency")
    axs[4].set_title("Grid Efficiency vs Year")
    axs[4].set_ylabel("Efficiency")
    axs[4].set_xlabel("Year")
    axs[4].legend()
    axs[4].grid(True)

    # Turn off unused subplot (bottom right)
    fig.delaxes(axs[5])

    plt.tight_layout()
    plt.show()

# Plot
plot_all_on_one_graph(df)
