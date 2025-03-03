import pandas as pd
import numpy as np
import time
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

# Load a local embedding model (Free & Open Source)
embedding_model = SentenceTransformer("sentence-transformers/all-MiniLM-L6-v2")

def generate_embedding(text):
    """Generate text embedding using Sentence-BERT."""
    if not text or pd.isna(text):
        return np.zeros(384)  # Return zero vector if text is missing (SBERT output size is 384)

    return embedding_model.encode(text)

def match_materials(database_file, material_takeoff_file, output_file):
    """ Matches materials from material takeoff to the database using Sentence-BERT embeddings. """

    print("\n🔄 Loading data...")
    df_db = pd.read_excel(database_file, sheet_name="Database")
    df_takeoff = pd.read_csv(material_takeoff_file)

    print("✅ Data loaded successfully!")

    # Combine relevant columns into a single descriptive string
    print("\n🔄 Preprocessing data...")
    df_db["Full_Description"] = df_db[["Material", "Sub-material", "ICE DB Name"]].fillna("").agg(" ".join, axis=1)
    df_takeoff["Full_Description"] = df_takeoff[["Category", "Material: Name", "Material: Description", "Description"]].fillna("").agg(" ".join, axis=1)
    print("✅ Data preprocessing complete!")

    # Generate embeddings for database materials
    print("\n🔄 Generating embeddings for database materials...")
    start_time = time.time()
    db_embeddings = np.vstack(df_db["Full_Description"].apply(generate_embedding))
    print(f"✅ Database embeddings generated! ⏱ Time taken: {round(time.time() - start_time, 2)} seconds")

    # Generate embeddings for material takeoff entries
    print("\n🔄 Generating embeddings for material takeoff materials...")
    start_time = time.time()
    takeoff_embeddings = np.vstack(df_takeoff["Full_Description"].apply(generate_embedding))
    print(f"✅ Takeoff embeddings generated! ⏱ Time taken: {round(time.time() - start_time, 2)} seconds")

    # Compute cosine similarity
    print("\n🔄 Computing similarity scores...")
    start_time = time.time()
    similarity_matrix = cosine_similarity(takeoff_embeddings, db_embeddings)
    print(f"✅ Similarity scores computed! ⏱ Time taken: {round(time.time() - start_time, 2)} seconds")

    # Find best match for each material in takeoff
    print("\n🔄 Finding best matches for each material...")
    best_match_indices = similarity_matrix.argmax(axis=1)
    
    # Assign Matched Unique ID, Average Embodied Carbon, and Density
    df_takeoff["Matched Unique ID"] = df_db.iloc[best_match_indices]["Unique ID"].values
    df_takeoff["Average Embodied Carbon"] = df_db.iloc[best_match_indices]["Average Embodied Carbon"].values
    df_takeoff["Density"] = df_db.iloc[best_match_indices]["Density"].values  # Include density from database
    
    print("✅ Matching complete!")

    # Save output file
    print("\n💾 Saving results to Excel...")
    df_takeoff.to_excel(output_file, sheet_name="Matched Takeoff", index=False)
    print(f"✅ Matching complete! Results saved in 📂 {output_file}")

# Example usage:
database_file = "C:/carbon_calculator/testing_material/database.xlsx"
material_takeoff_file = "C:/carbon_calculator/testing_material/material_takeoff.csv"
output_file = "C:/carbon_calculator/testing_material/matched_material_takeoff.xlsx"

match_materials(database_file, material_takeoff_file, output_file)