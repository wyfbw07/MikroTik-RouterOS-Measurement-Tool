import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

pre_path = 'Results_Excel/'
# Load the Excel file
file_path = pre_path+'5-2.xlsx'
file_path2 = pre_path+'10-1.xlsx'
file_path3 = pre_path+'20-1.xlsx'
file_path4 = pre_path+'30-1.xlsx'
file_path5 = pre_path+'40-1.xlsx'
file_path6 = pre_path+'50-1.xlsx'
data = pd.ExcelFile(file_path)
data2 = pd.ExcelFile(file_path2)
data3 = pd.ExcelFile(file_path3)
data4 = pd.ExcelFile(file_path4)
data5 = pd.ExcelFile(file_path5)
data6 = pd.ExcelFile(file_path6)

# Load the 'Measurements' sheet
measurements_data = data.parse('Set_1')
measurements_data2 = data2.parse('Set_1')
measurements_data3 = data3.parse('Set_1')
measurements_data4 = data4.parse('Set_1')
measurements_data5 = data5.parse('Set_1')
measurements_data6 = data6.parse('Set_1')

# Extract relevant columns
sequence_numbers = measurements_data['seq_no']
latency = measurements_data['delay (ms)']
error = measurements_data['err (ms)']

sequence_numbers2 = measurements_data2['seq_no']
latency2 = measurements_data2['delay (ms)']
error2 = measurements_data2['err (ms)']

sequence_numbers3 = measurements_data3['seq_no']
latency3 = measurements_data3['delay (ms)']
error3 = measurements_data3['err (ms)']

sequence_numbers4 = measurements_data4['seq_no']
latency4 = measurements_data4['delay (ms)']
error4 = measurements_data4['err (ms)']

sequence_numbers5 = measurements_data5['seq_no']
latency5 = measurements_data5['delay (ms)']
error5 = measurements_data5['err (ms)']

sequence_numbers6 = measurements_data6['seq_no']
latency6 = measurements_data6['delay (ms)']
error6 = measurements_data6['err (ms)']

# Define gradient colors from shallow to deep
colors = plt.cm.Reds(np.linspace(0.15, 0.8, 6))  # Adjust the colormap and range as needed

# Plot Line Chart with gradient colors
plt.figure(figsize=(10, 6))
# plt.plot(sequence_numbers, latency, label='5m', color=colors[0])
plt.plot(sequence_numbers2, latency2, label='10m', color=colors[1])
plt.plot(sequence_numbers3, latency3, label='20m', color=colors[2])
plt.plot(sequence_numbers4, latency4, label='30m', color=colors[3])
plt.plot(sequence_numbers5, latency5, label='40m', color=colors[4])
plt.plot(sequence_numbers6, latency6, label='50m', color=colors[5])
plt.title('Latency Over Sequence Numbers')
plt.xlabel('Sequence Number')
plt.ylabel('Latency (ms)')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()

# Plot Histogram
bin_num = 10
plt.figure(figsize=(10, 6))
# plt.hist(latency, bins=bin_num, alpha=0.7, label='5m', color=colors[0])
plt.hist(latency2, bins=bin_num, alpha=0.7, label='10m', color=colors[1])
plt.hist(latency3, bins=bin_num, alpha=0.7, label='20m', color=colors[2])
plt.hist(latency4, bins=bin_num, alpha=0.7, label='30m', color=colors[3])
plt.hist(latency5, bins=bin_num, alpha=0.7, label='40m', color=colors[4])
plt.hist(latency6, bins=bin_num, alpha=0.7, label='50m', color=colors[5])
plt.title('Frequency of Latencies over different distances')
plt.xlabel('Time (ms)')
plt.ylabel('Frequency')
plt.legend()
plt.tight_layout()
plt.show()