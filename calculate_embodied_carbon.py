import pandas as pd
import re

def calculate_embodied_carbon(final_excel_file):
    """Calculates the total embodied carbon and weighted DQI from the final matched material takeoff sheet."""

    print("\n🔄 Loading final matched material takeoff sheet...")
    df = pd.read_excel(final_excel_file, sheet_name="Matched Takeoff")
    print("✅ File loaded successfully!")

    # Remove "m³" from Material: Volume and convert to float
    df["Material: Volume"] = df["Material: Volume"].astype(str).apply(lambda x: re.sub(r"[^0-9.]", "", x)).astype(float)

    # Calculate embodied carbon for each row: Volume (m³) × Density (kg/m³) × Average Embodied Carbon (kg CO2e/kg)
    df["Row Embodied Carbon (kg CO2e)"] = df["Material: Volume"] * df["Density"] * df["Average Embodied Carbon"]

    # Total Embodied Carbon
    total_embodied_carbon = df["Row Embodied Carbon (kg CO2e)"].sum()

    # Compute Weighted DQI: sum(DQI * Embodied Carbon) / sum(Embodied Carbon)
    df["DQI Contribution"] = df["DQI"] * df["Row Embodied Carbon (kg CO2e)"]
    weighted_dqi = df["DQI Contribution"].sum() / total_embodied_carbon

    # Print results
    print("\n🔹 Embodied Carbon Calculations:")
    print(f"🌍 Total Embodied Carbon: {total_embodied_carbon:,.2f} kg CO2e")
    print(f"📊 Weighted DQI: {100*weighted_dqi:.2f}%\n")

    return total_embodied_carbon, weighted_dqi
