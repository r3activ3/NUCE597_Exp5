import pandas as pd
import plotly.graph_objects as go

# Load the Excel file
file_path = r"C:\Users\17244\OneDrive\Grad School\NUCE 597\Exp5\Data.xlsx"
data = pd.read_excel(file_path)

# Function to create the plots
def create_plots(data):
    plots = []
    #channels = data['Channel']
    energy = data['Energy']
    
    columns = data.columns[2:]
    
    for col in columns:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=energy, y=data[col], mode='lines+markers', name=col))
        fig.update_layout(
            title=f"{col} vs Energy (keV)",
            xaxis_title="Energy (keV)",
            yaxis_title=col,
            hovermode="x unified"
        )
        plots.append(fig)
    
    return plots

# Generate the plots
plots = create_plots(data)

# Display the plots (example for the first few plots)
for plot in plots:
    plot.show()