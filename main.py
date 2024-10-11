import paramiko
import time
import pandas as pd
from openpyxl import Workbook

def execute_command(ssh, command):
    stdin, stdout, stderr = ssh.exec_command(command)
    output = stdout.read().decode().strip()
    error = stderr.read().decode().strip()
    
    return output if output else f"Error: {error}"

def main():
    # Create an SSH client
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    # Connect to the MikroTik router
    try:
        ssh.connect('192.168.2.2', username='admin', password='rpi')
    except paramiko.SSHException as e:
        print(f"Error connecting to the host: {e}")
        return

    # Define the commands to execute
    commands = [
        'connected', 'frequency', 'remote-address', 'tx-mcs', 'tx-phy-rate', 
        'signal', 'rssi', 'tx-sector', 'tx-sector-info', 'distance', 'tx-packet-error-rate'
    ]
    
    # Initialize a list to store the data for each iteration
    data = []

    try:
        while True:
            # Collect results
            results = [execute_command(ssh, f'/interface w60g {{:put ([monitor wlan60-1 once as-value ]->"{cmd}")}}') for cmd in commands]
            
            # Append results to data list
            data.append(results)
            
            # Convert data into a DataFrame for easier management
            df = pd.DataFrame(data, columns=commands)
            
            # Save the DataFrame to an Excel file
            df.to_excel('mikrotik_data.xlsx', index=False)
            
            # Wait for 1 second before the next iteration
            time.sleep(1)
    except KeyboardInterrupt:
        print("Stopped the script.")
    finally:
        ssh.close()

if __name__ == "__main__":
    main()
