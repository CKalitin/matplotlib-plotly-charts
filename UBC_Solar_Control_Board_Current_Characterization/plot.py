import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import os
import sys

# Run in terminal with `python plot.py <filename>`

# Files:
# TEK0008.CSV = First run recording startup, 5 second total time
# TEK0012.CSV = Startup with 100 ms div size, for 1s total recording time
# TEK0015.CSV = Full startup sequence with 1s div size for 10s total recording time
#

def smooth_data(data, window_size=11):
    """
    Apply moving average smoothing to data to reduce noise.
    
    Args:
        data (array): Input data to smooth
        window_size (int): Size of the smoothing window
    
    Returns:
        array: Smoothed data
    """
    if len(data) < window_size:
        return data
    
    # Pad the data at the edges to handle boundary effects
    pad_width = window_size // 2
    padded_data = np.pad(data, pad_width, mode='edge')
    
    # Apply moving average filter, 11 points
    smoothed = np.convolve(padded_data, np.ones(window_size)/window_size, mode='same')
    
    # Remove padding
    return smoothed[pad_width:-pad_width]

def read_tek_csv(filename):
    """
    Read a Tektronix CSV file and extract time and voltage data.
    
    Args:
        filename (str): Path to the CSV file
        
    Returns:
        tuple: (time_data, voltage_data) as numpy arrays
    """
    # Read the CSV file, skipping the header information
    # The actual data starts after the metadata rows
    data = []
    time_data = []
    voltage_data = []
    
    with open(filename, 'r') as file:
        lines = file.readlines()
        
    # Find where the actual data starts (rows with time and voltage values)
    for line in lines:
        parts = line.strip().split(',')
        if len(parts) >= 5:
            try:
                # Try to parse the 4th and 5th columns as time and voltage
                time_val = float(parts[3])
                voltage_val = float(parts[4])
                time_data.append(time_val)
                voltage_data.append(voltage_val)
            except (ValueError, IndexError):
                # Skip rows that can't be parsed (metadata rows)
                continue
    
    return np.array(time_data), np.array(voltage_data)

def plot_voltage_and_current(filename, sensitivity_mv_per_a=200):
    """
    Create time vs voltage and time vs current plots for a given CSV file.
    
    Args:
        filename (str): Path to the CSV file
        sensitivity_mv_per_a (float): Sensitivity in mV/A (default: 200)
    """
    # Read the data
    time, voltage = read_tek_csv(filename)
    
    # Apply voltage offset (subtract 1.65V)
    voltage = voltage - 1.65
    
    # Set any negative voltage values to the default 0 V
    voltage[voltage < 0] = 0
    
    # Apply smoothing to reduce noise
    voltage_smooth = smooth_data(voltage, window_size=11)
    
    # Convert sensitivity from mV/A to V/A
    sensitivity_v_per_a = sensitivity_mv_per_a / 1000
    
    # Calculate current using I = V / sensitivity (using smoothed voltage)
    current = voltage / sensitivity_v_per_a
    
    # Apply smoothing to current as well
    current_smooth = smooth_data(current, window_size=11)
    
    # Get base filename for saving
    base_filename = os.path.splitext(os.path.basename(filename))[0]
    
    # Create voltage plot
    fig1, ax1 = plt.subplots(1, 1, figsize=(12, 6))
    ax1.plot(time, voltage_smooth, 'b-', linewidth=1, label='Smoothed')
    ax1.plot(time, voltage, 'b-', linewidth=0.3, alpha=0.7, label='Raw')
    ax1.set_xlabel('Time (s)')
    ax1.set_ylabel('Voltage (V)')
    ax1.set_title(f'Voltage vs Time - {os.path.basename(filename)} (Offset: -1.65V)')
    ax1.grid(True, alpha=0.3)
    ax1.legend()
    
    # Save voltage plot
    voltage_filename = f"{base_filename}_voltage.png"
    plt.tight_layout()
    plt.savefig(voltage_filename, dpi=300, bbox_inches='tight')
    print(f"Voltage plot saved as: {voltage_filename}")
    plt.show()
    
    # Create current plot
    fig2, ax2 = plt.subplots(1, 1, figsize=(12, 6))
    ax2.plot(time, current_smooth, 'r-', linewidth=1, label='Smoothed')
    ax2.plot(time, current, 'r-', linewidth=0.3, alpha=0.7, label='Raw')
    ax2.set_xlabel('Time (s)')
    ax2.set_ylabel('Current (A)')
    ax2.set_title(f'Current vs Time (Sensitivity: {sensitivity_mv_per_a} mV/A, Smoothed)')
    ax2.grid(True, alpha=0.3)
    ax2.legend()
    
    # Save current plot
    current_filename = f"{base_filename}_current.png"
    plt.tight_layout()
    plt.savefig(current_filename, dpi=300, bbox_inches='tight')
    print(f"Current plot saved as: {current_filename}")
    plt.show()
    
    # Print some statistics
    print(f"\nFile: {filename}")
    print(f"Time range: {time.min():.6f} s to {time.max():.6f} s")
    print(f"Raw voltage range: {voltage.min():.4f} V to {voltage.max():.4f} V")
    print(f"Smoothed voltage range: {voltage_smooth.min():.4f} V to {voltage_smooth.max():.4f} V")
    print(f"Raw current range: {current.min():.4f} A to {current.max():.4f} A")
    print(f"Smoothed current range: {current_smooth.min():.4f} A to {current_smooth.max():.4f} A")
    print(f"Data points: {len(time)}")
    print(f"Smoothing window: 11 points")

def main():
    """
    Main function to handle command line arguments and file selection.
    """
    if len(sys.argv) > 1:
        # If filename is provided as command line argument
        filename = sys.argv[1]
    else:
        # Interactive file selection
        print("Available CSV files:")
        csv_files = [f for f in os.listdir('.') if f.endswith('.CSV')]
        csv_files.sort()
        
        for i, file in enumerate(csv_files):
            print(f"{i}: {file}")
        
        try:
            choice = int(input("\nEnter the number of the file to plot: "))
            filename = csv_files[choice]
        except (ValueError, IndexError):
            print("Invalid selection. Exiting.")
            return
    
    # Check if file exists
    if not os.path.exists(filename):
        print(f"Error: File '{filename}' not found.")
        return
    
    # Use fixed sensitivity of 200 mV/A
    sensitivity = 200
    
    # Create the plots
    plot_voltage_and_current(filename, sensitivity)

if __name__ == "__main__":
    main()