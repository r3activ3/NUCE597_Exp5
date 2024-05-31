import pandas as pd
import plotly.graph_objects as go
import numpy as np

# Load the Excel file
file_path = r"C:\Users\17244\OneDrive\Grad School\NUCE 597\Exp5\Data.xlsx"
data = pd.read_excel(file_path)

# Constants for correction factor (assuming no container effects for simplicity)
mu = 1.0  # Mass attenuation coefficient
rho = 1.0  # Density
t = 1.0  # Thickness

# Known enrichment and target energy
known_enrichment = 15.41  # %
target_energy = 185.7  # keV

# Find the row with the energy closest to 185.7 keV
closest_energy_row = data.iloc[(data['Energy'] - target_energy).abs().argsort()[:1]]

# Extract the counts and uncertainties for the enriched sample at 15.41% from the closest energy row
counts = closest_energy_row['Enriched 15.41% Counts'].values[0]
uncertainty_counts = closest_energy_row['Enriched 15.41% Uncertainty'].values[0]

# Calculate the net count rate and its uncertainty
net_count_rate = counts / 900  # Counts per 900 seconds
net_count_rate_uncertainty = uncertainty_counts

# Calculate the correction factor (CF) and its uncertainty
CF = np.exp(mu * rho * t)
CF_uncertainty = CF * np.sqrt((t * rho * mu)**2 + (t * mu * rho)**2 + (mu * rho * t)**2)

# Calculate the calibration constant (K) and its uncertainty
K = known_enrichment / (net_count_rate * CF)
K_uncertainty = K * np.sqrt((net_count_rate_uncertainty / net_count_rate)**2 + (CF_uncertainty / CF)**2)

# Function to apply the calibration constant to estimate enrichment
def calculate_enrichment(count_rate, K, CF, K_uncertainty, CF_uncertainty):
    enrichment = K * count_rate * CF
    uncertainty_enrichment = enrichment * np.sqrt((K_uncertainty / K)**2 + (net_count_rate_uncertainty / count_rate)**2 + (CF_uncertainty / CF)**2)
    return enrichment, uncertainty_enrichment

# Apply the calibration constant to all samples
data['Calculated Enrichment'] = data['Enriched 15.41% Counts'].apply(lambda x: calculate_enrichment(x / 900, K, CF, K_uncertainty, CF_uncertainty)[0])
data['Calculated Enrichment Uncertainty'] = data['Enriched 15.41% Counts'].apply(lambda x: calculate_enrichment(x / 900, K, CF, K_uncertainty, CF_uncertainty)[1])

# Generate the plot for energy vs count rate
def plot_energy_vs_count_rate(data):
    energy = data['Energy']
    columns = ['Americium 241 Counts', 'Background Counts', 'Natural 0.07% Counts', 'Natural slug Counts', 'Depleted U 0.02% Counts', 
               'Depleted U 0.5% Counts', 'Enriched 15.41% Counts', 'HEU Counts', 'Fuel Pellet 1 Counts', 'Fuel Pellet 2 Counts']
    
    fig = go.Figure()
    for col in columns:
        fig.add_trace(go.Scatter(x=energy, y=data[col], mode='lines+markers', name=col))
    
    fig.update_layout(
        title="Energy vs Counts for All Samples",
        xaxis_title="Energy (keV)",
        yaxis_title="Count Rate",
        hovermode="x unified"
    )
    
    return fig

# Function to create the combined plot for enriched samples within a specified energy range
def create_combined_plot_in_range(data, energy_range, title):
    energy = data['Energy']
    columns = ['Natural 0.07% Counts', 'Natural slug Counts', 'Depleted U 0.02% Counts', 
               'Depleted U 0.5% Counts', 'Enriched 15.41% Counts', 'HEU Counts', 'Fuel Pellet 1 Counts', 'Fuel Pellet 2 Counts']
    
    fig = go.Figure()
    for col in columns:
        fig.add_trace(go.Scatter(x=energy[energy_range], y=data[col][energy_range], mode='lines+markers', name=col))
    
    fig.update_layout(
        title=title,
        xaxis_title="Energy (keV)",
        yaxis_title="Counts",
        hovermode="x unified"
    )
    
    return fig

# Generate the combined plots for the specified energy ranges
energy_range_184_187 = (data['Energy'] >= 184) & (data['Energy'] <= 187)
energy_range_88_102 = (data['Energy'] >= 88) & (data['Energy'] <= 102)

combined_plot_184_187 = create_combined_plot_in_range(data, energy_range_184_187, "Enriched Samples vs Energy EMP Region")
combined_plot_88_102 = create_combined_plot_in_range(data, energy_range_88_102, "Enriched Samples vs Energy MGAU Region")

# Display the plots
energy_vs_count_rate_plot = plot_energy_vs_count_rate(data)
energy_vs_count_rate_plot.show()
combined_plot_184_187.show()
combined_plot_88_102.show()