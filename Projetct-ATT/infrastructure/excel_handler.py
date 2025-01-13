import pandas as pd

def load_excel(file_path):
    return pd.read_excel(file_path, engine='openpyxl')

def save_excel(file_path, df):
    df.to_excel(file_path, index=False, engine='openpyxl')
