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
    time.sleep(0.1)
    s.send('string temp_string;'.encode())
    time.sleep(0.1)
    s.send('sprintf(temp_string, "%f", temp_float);'.encode())
    time.sleep(0.1)
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
    "Bus107A": "V107A_rms"
    
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

    for load_name, config in loads.items():
        load_val = base_load * (1 + growth_rate) ** (year - 1)
        change_slider_val(s, config["slider"], load_val)
        time.sleep(wait_time)

        p = get_meter_val(s, config["P"])
        q = get_meter_val(s, config["Q"])

        year_row[f"{load_name}_Load"] = load_val
        year_row[f"{load_name}_P"] = p
        year_row[f"{load_name}_Q"] = q

        print(f"[Year {year}] {load_name} → Load: {load_val:.2f}, P: {p:.3f}, Q: {q:.3f}")

    for label, meter in bus_voltage_meters.items():
        voltage = get_meter_val(s, meter)
        year_row[label] = voltage
        print(f"[Year {year}] {label}: {voltage:.3f}")

    all_data.append(year_row)

#plots 
def plot_all_on_one_graph(df):
    fig, axs = plt.subplots(3, 1, figsize=(12, 12), sharex=True)

    # Real Power (P) 
    p_cols = [col for col in df.columns if col.endswith('_P')]
    for col in p_cols:
        axs[0].plot(df["Year"], df[col], marker='o', label=col)
    axs[0].set_title("Real Power (P) vs Year")
    axs[0].set_ylabel("Power (P)")
    axs[0].legend()
    axs[0].grid(True)

    #Reactive Power (Q) 
    q_cols = [col for col in df.columns if col.endswith('_Q')]
    for col in q_cols:
        axs[1].plot(df["Year"], df[col], marker='o', label=col)
    axs[1].set_title("Reactive Power (Q) vs Year")
    axs[1].set_ylabel("Reactive Power (Q)")
    axs[1].legend()
    axs[1].grid(True)

    # Bus Voltages 
    bus_cols = [col for col in df.columns if col.startswith("Bus")]
    for col in bus_cols:
        axs[2].plot(df["Year"], df[col], marker='o', label=col)
    axs[2].set_title("Bus Voltages vs Year")
    axs[2].set_xlabel("Year")
    axs[2].set_ylabel("Voltage (p.u.)")
    axs[2].legend()
    axs[2].grid(True)

    plt.tight_layout()
    plt.show()

#Excel
save_path = "C:/Users/CESAC/Desktop/Bijay/Data/load_growth_results.xlsx"
df = pd.DataFrame(all_data)
df.to_excel(save_path, index=False)
print("\nExcel file saved")

#plot
plot_all_on_one_graph(df)
