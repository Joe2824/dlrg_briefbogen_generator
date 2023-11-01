#!/usr/bin/env python
import warnings
import os
import pandas as pd
import argparse

# Parse command-line arguments
parser = argparse.ArgumentParser(description="Convert Excel data to Pickle format.")
parser.add_argument("excel_file_path", help="Path to the Excel file")
args = parser.parse_args()

# Suppress openpyxl warning
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Get the current working directory
current_directory = os.getcwd()
parent_directory = os.path.dirname(current_directory)
resources_folder = os.path.join(parent_directory, 'resources')

# Load Excel data
excel_file_path = args.excel_file_path
df_dict = pd.read_excel(excel_file_path, sheet_name=None, engine="openpyxl")
allgemein_df = df_dict.get('Allgemeine Daten')
vorstand_df = df_dict.get('Vorstand')
jugend_df = df_dict.get('Jugend')

# Define output paths
output_dir = os.path.join('src', 'pkl')
os.makedirs(output_dir, exist_ok=True)

# Save DataFrames to Pickle format
allgemein_df.to_pickle(os.path.join(output_dir, 'general.pkl'), compression='xz')
vorstand_df.to_pickle(os.path.join(output_dir, 'vorstand.pkl'), compression='xz')
jugend_df.to_pickle(os.path.join(output_dir, 'jugend.pkl'), compression='xz')

print("Data saved to Pickle files.")
