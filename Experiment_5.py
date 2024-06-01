import pandas as pd
import plotly.graph_objects as go
import numpy as np

# Load the Excel file
file_path = r"C:\Users\jerem\OneDrive\Grad School\NUCE 597\Exp5\Data.xlsx"
data = pd.read_excel(file_path)

samples = ['Natural 0.07% Counts', 'Natural slug Counts', 'Depleted U 0.02% Counts', 
           'Depleted U 0.5% Counts', 'Enriched 15.41% Counts', 'HEU Counts', 'Fuel Pellet 1 Counts', 'Fuel Pellet 2 Counts']

# Known enrichment and target energy
known_enrichment = 15.41  # %
target_energy = 185.49307  # keV

# Find the row with the energy closest to target energy
closest_energy_row = data.iloc[(data['Energy'] - target_energy).abs().argsort()[:1]]

# Extract the counts and uncertainties for the enriched sample at 15.41% from the closest energy row
counts = closest_energy_row['Enriched 15.41% Counts'].values[0]
uncertainty_counts = closest_energy_row['Enriched 15.41% Counts Uncertainty'].values[0]

# Calculate the net count rate and its uncertainty
net_count_rate = counts / 1800  # Counts per 1800 seconds
net_count_rate_uncertainty = uncertainty_counts


# Calculate the calibration constant (K) and its uncertainty
a = 1/((known_enrichment/net_count_rate)*((net_count_rate_uncertainty/net_count_rate)**2))
b = 1/(((known_enrichment/net_count_rate)**2) * ((net_count_rate_uncertainty/net_count_rate)**2))
K = a/b
K_uncertainty = np.sqrt(1/(1/(((known_enrichment/net_count_rate)**2)*((0.01/known_enrichment)**2))))


# Function to apply the calibration constant to estimate enrichment
def calculate_enrichment(count_rate, K,CR_uncertainty):
    enrichment = K * count_rate
    uncertainty_enrichment = enrichment * (CR_uncertainty/count_rate)
    return enrichment, uncertainty_enrichment

for sample in samples:
    count_rate =  (closest_energy_row[sample].values[0])/1800
    count_rate_uncertainty = closest_energy_row[f'{sample} Uncertainty'].values[0]
    enrichment, uncertainty_enrichment= calculate_enrichment(count_rate, K, count_rate_uncertainty)
    print(sample, enrichment, uncertainty_enrichment)

# Generate the plot for energy vs count rate
def plot_energy_vs_count_rate(data):
    energy = data['Energy']
    columns = ['Americium 241 Counts', 'Background Counts', 'Natural 0.07% Counts', 'Natural slug Counts', 'Depleted U 0.02% Counts', 
               'Depleted U 0.5% Counts', 'Enriched 15.41% Counts', 'HEU Counts', 'Fuel Pellet 1 Counts', 'Fuel Pellet 2 Counts']
    
    fig = go.Figure()
    for col in columns:
        fig.add_trace(go.Scatter(x=energy, y=data[col], mode='lines', name=col))
    
    fig.update_layout(
        title="Energy vs Counts for All Samples",
        xaxis_title="Energy (keV)",
        yaxis_title="Counts",
        hovermode="x unified",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.3,
            xanchor="center",
            x=0.5
        )
    )
    
    return fig

# Function to create the combined plot for enriched samples within a specified energy range
def create_combined_plot_in_range(data, energy_range, title):
    energy = data['Energy']
    columns = ['Natural 0.07% Counts', 'Natural slug Counts', 'Depleted U 0.02% Counts', 
               'Depleted U 0.5% Counts', 'Enriched 15.41% Counts', 'HEU Counts', 'Fuel Pellet 1 Counts', 'Fuel Pellet 2 Counts']
    
    fig = go.Figure()
    for col in columns:
        fig.add_trace(go.Scatter(x=energy[energy_range], y=data[col][energy_range], mode='lines', name=col))
    
    fig.update_layout(
        title=title,
        xaxis_title="Energy (keV)",
        yaxis_title="Counts",
        hovermode="x unified",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.3,
            xanchor="center",
            x=0.5
        )
    )
    
    return fig

# Generate the combined plots for the specified energy ranges
energy_range_184_187 = (data['Energy'] >= 184) & (data['Energy'] <= 187)
energy_range_88_102 = (data['Energy'] >= 88) & (data['Energy'] <= 102)

combined_plot_184_187 = create_combined_plot_in_range(data, energy_range_184_187, "Enriched Samples vs Energy EMP Region")
combined_plot_88_102 = create_combined_plot_in_range(data, energy_range_88_102, "Enriched Samples vs Energy MGAU Region")

# Display the plots
energy_vs_count_rate_plot = plot_energy_vs_count_rate(data)

combined_plot_184_187.show()
combined_plot_88_102.show()
energy_vs_count_rate_plot.show()