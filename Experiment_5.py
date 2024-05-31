import pandas as pd
import plotly.graph_objects as go

# Load the Excel file
file_path = r"C:\Users\17244\OneDrive\Grad School\NUCE 597\Exp5\Data.xlsx"
data = pd.read_excel(file_path)

# Function to create the individual plots with background energy removed
def create_plots(data):
    plots = []
    energy = data['Energy']
    background = data['Background']
    
    columns = data.columns[2:]
    
    for col in columns:
        if col != 'Background':
            adjusted_data = data[col] - background
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=energy, y=adjusted_data, mode='lines+markers', name=col))
            fig.update_layout(
                title=f"{col} vs Energy (keV) with Background Removed",
                xaxis_title="Energy (keV)",
                yaxis_title=f"{col} (Adjusted)",
                hovermode="x unified"
            )
            plots.append(fig)
    
    return plots

# Function to create a combined plot for all enriched samples
def create_combined_plot(data, energy_range=None):
    energy = data['Energy']
    background = data['Background']
    
    enriched_columns = ['Natural 0.07%', 'Natural slug', 'Depleted U 0.02%', 'Depleted U 0.5%', 'Enriched 15.41%', 'HEU', 'Fuel Pellet 1', 'Fuel Pellet 2']
    
    fig = go.Figure()
    
    for col in enriched_columns:
        adjusted_data = data[col] - background
        if energy_range:
            mask = (energy >= energy_range[0]) & (energy <= energy_range[1])
            fig.add_trace(go.Scatter(x=energy[mask], y=adjusted_data[mask], mode='lines+markers', name=col))
        else:
            fig.add_trace(go.Scatter(x=energy, y=adjusted_data, mode='lines+markers', name=col))
    
    fig.update_layout(
        title="Enriched Samples vs Energy (keV) with Background Removed",
        xaxis_title="Energy (keV)",
        yaxis_title="Adjusted Counts",
        hovermode="x unified"
    )
    
    return fig

# Generate the individual plots
individual_plots = create_plots(data)

# Generate the combined plot for enriched samples
combined_plot = create_combined_plot(data)

# Generate the combined plots for the specified energy ranges
combined_plot_184_187 = create_combined_plot(data, energy_range=(184, 187))
combined_plot_88_102 = create_combined_plot(data, energy_range=(88, 102))

# Display the individual plots
for plot in individual_plots:
    plot.show()

# Display the combined plot
combined_plot.show()

# Display the combined plots for the specified energy ranges
combined_plot_184_187.show()
combined_plot_88_102.show()