import pandas as pd
import re

def summarize_embodied_carbon(final_excel_file, database_file):
    """Summarizes the embodied carbon contributions for each unique material in the takeoff sheet."""

    print("\n🔄 Loading final matched material takeoff sheet...")
    df_takeoff = pd.read_excel(final_excel_file, sheet_name="Matched Takeoff")
    df_db = pd.read_excel(database_file, sheet_name="Database")  # Load database for material names
    print("✅ Files loaded successfully!")

    # Remove "m³" from Material: Volume and convert to float
    df_takeoff["Material: Volume"] = df_takeoff["Material: Volume"].astype(str).apply(lambda x: re.sub(r"[^0-9.]", "", x)).astype(float)

    # Calculate embodied carbon per row: Volume (m³) × Density (kg/m³) × Average Embodied Carbon (kg CO2e/kg)
    df_takeoff["Row Embodied Carbon (kg CO2e)"] = df_takeoff["Material: Volume"] * df_takeoff["Density"] * df_takeoff["Average Embodied Carbon"]

    # Compute total embodied carbon for each unique material (Unique ID)
    material_summary = df_takeoff.groupby("Matched Unique ID")["Row Embodied Carbon (kg CO2e)"].sum().reset_index()

    # Compute total embodied carbon for the entire building
    total_embodied_carbon = material_summary["Row Embodied Carbon (kg CO2e)"].sum()

    # Calculate percentage contribution of each material
    material_summary["% Embodied Carbon"] = (material_summary["Row Embodied Carbon (kg CO2e)"] / total_embodied_carbon) * 100

    # Merge with database to get Material and Sub-material names
    material_summary = material_summary.merge(df_db[["Unique ID", "Material", "Sub-material"]], left_on="Matched Unique ID", right_on="Unique ID", how="left")

    # Select final columns in the required order
    material_summary = material_summary[["Matched Unique ID", "Material", "Sub-material", "Row Embodied Carbon (kg CO2e)", "% Embodied Carbon"]]

    # Print results in console
    print("\n🔹 Embodied Carbon Contribution by Material:")
    print(material_summary.to_string(index=False))

    return material_summary