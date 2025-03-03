import pandas as pd
import re

def summarize_embodied_carbon_by_component(final_excel_file):
    """Summarizes embodied carbon contributions by building component (Walls, Floors, Roof, etc.)."""

    print("\n🔄 Loading final matched material takeoff sheet...")
    df = pd.read_excel(final_excel_file, sheet_name="Matched Takeoff")
    print("✅ File loaded successfully!")

    # Remove "m³" from Material: Volume and convert to float
    df["Material: Volume"] = df["Material: Volume"].astype(str).apply(lambda x: re.sub(r"[^0-9.]", "", x)).astype(float)

    # Calculate embodied carbon per row: Volume (m³) × Density (kg/m³) × Average Embodied Carbon (kg CO2e/kg)
    df["Row Embodied Carbon (kg CO2e)"] = df["Material: Volume"] * df["Density"] * df["Average Embodied Carbon"]

    # Compute total embodied carbon for each component (Category)
    component_summary = df.groupby("Category")["Row Embodied Carbon (kg CO2e)"].sum().reset_index()

    # Compute total embodied carbon for the entire building
    total_embodied_carbon = component_summary["Row Embodied Carbon (kg CO2e)"].sum()

    # Calculate percentage contribution of each component
    component_summary["% Embodied Carbon"] = (component_summary["Row Embodied Carbon (kg CO2e)"] / total_embodied_carbon) * 100

    # Print results in console
    print("\n🔹 Embodied Carbon Contribution by Building Component:")
    print(component_summary.to_string(index=False))

    return component_summary