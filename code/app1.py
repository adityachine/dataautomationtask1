import pandas as pd
import logging
import os

# === Step 1: Setup Logging ===
logging.basicConfig(
    filename='automation_log.txt',
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

# === Step 2: Load CSV Data ===
def load_data(filepath):
    try:
        df = pd.read_csv(filepath)
        logging.info("CSV file loaded successfully.")
        return df
    except FileNotFoundError:
        logging.error(f"File not found: {filepath}")
        raise
    except Exception as e:
        logging.error(f"Error loading data: {e}")
        raise

# === Step 3: Clean and Preprocess Data ===
def clean_data(df):
    try:
        # Convert boolean to strings if needed
        df['Weekend'] = df['Weekend'].astype(str)
        df['Revenue'] = df['Revenue'].astype(str)

        # Fill missing numerical values with 0
        num_cols = df.select_dtypes(include=['float64', 'int64']).columns
        df[num_cols] = df[num_cols].fillna(0)

        logging.info("Data cleaned and preprocessed successfully.")
        return df
    except Exception as e:
        logging.error(f"Error in data cleaning: {e}")
        raise

# === Step 4: Create Pivot Table ===
def create_pivot_table(df):
    try:
        pivot = pd.pivot_table(
            df,
            values='PageValues',
            index='VisitorType',
            columns='Weekend',
            aggfunc='mean',
            fill_value=0
        )
        logging.info("Pivot table created successfully.")
        return pivot
    except Exception as e:
        logging.error(f"Error creating pivot table: {e}")
        raise

# === Step 5: Save Pivot Table ===
def save_output(pivot_df, output_path):
    try:
        pivot_df.to_csv(output_path)
        logging.info(f"Pivot table exported to {output_path}")
    except Exception as e:
        logging.error(f"Error saving output: {e}")
        raise

# === Main Workflow ===
def automate_workflow():
    input_path = 'online_shoppers_intention.csv'  # Ensure this path is correct
    output_path = 'pivot_output.csv'

    print("🔄 Starting automation...")
    df = load_data(input_path)
    df = clean_data(df)
    pivot = create_pivot_table(df)
    save_output(pivot, output_path)
    print("✅ Automation complete. Output saved as:", output_path)

if __name__ == "__main__":
    automate_workflow()
