import matplotlib.pyplot as plt
import numpy as np

# Data from the table
rows = [8, 16, 32, 64, 128]
columns = [8, 16, 32, 64, 128]

# Mean times for increasing rows, with columns fixed
compute_time_fixed_col_8 = [101.1, 177.0, 438.2, 659.6, 1363.6]
compute_time_fixed_col_16 = [196.0, 349.8, 690.6, 1392.5, 2711.4]

# Mean times for increasing columns, with rows fixed
compute_time_fixed_row_8 = [101.1, 196.0, 396.0, 688.3, 1416.3]
compute_time_fixed_row_16 = [177.0, 349.8, 689.4, 1376.7, 3488.1]

# Plotting
plt.figure(figsize=(12, 6))

# Subplot 1: Increasing Rows, Fixed Columns
plt.subplot(1, 2, 1)
plt.plot(rows, compute_time_fixed_col_8, 'o-', label="COL=8", color='blue')
plt.plot(rows, compute_time_fixed_col_16, 'o-', label="COL=16", color='orange')
plt.xlabel('Number of Rows')
plt.ylabel('Compute Time (ms)')
plt.title('Compute Time vs Rows (Fixed COL)\nLess, but nearing double (for larger values) compute time for each double row')
plt.legend()
plt.grid(True)

# Subplot 2: Increasing Columns, Fixed Rows
plt.subplot(1, 2, 2)
plt.plot(columns, compute_time_fixed_row_8, 'o-', label="ROW=8", color='green')
plt.plot(columns, compute_time_fixed_row_16, 'o-', label="ROW=16", color='red')
plt.xlabel('Number of Columns')
plt.ylabel('Compute Time (ms)')
plt.title('Compute Time vs Columns (Fixed ROW)\nMore than double compute time for double columns')
plt.legend()
plt.grid(True)

# Show the plots
plt.tight_layout()
plt.show()
