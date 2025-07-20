# data_processor.py
import pandas as pd
import numpy as np
from config import Config

class DataProcessor:
    def __init__(self):
        self.config = Config()
    
    def clean_data(self, dataframes):
        """データクリーニング処理"""
        if not dataframes:
            return None
        
        final_df = pd.concat(dataframes, ignore_index=True)
        
        for col in ['物件名', '向き']:
            if col in final_df.columns:
                final_df[col] = final_df[col].astype(str)
        
        df_sorted = final_df.sort_values(by="物件名", ascending=True)
        
        # 不要な行を削除
        df_sorted = df_sorted[~df_sorted["物件名"].str.contains("築", na=False)]
        df_sorted = df_sorted[~df_sorted["物件名"].str.contains("号室", na=False)]
        df_sorted = df_sorted[~df_sorted["向き"].str.contains("-", na=False)]
        
        # 重複削除
        columns = ['賃料(円)', '管理費', '間取り', '専有面積(㎡)', '向き']
        df_sorted = self.remove_duplicates_multiple_sorts(df_sorted, columns)
        
        # NAを削除
        df_sorted = df_sorted.replace('', pd.NA).dropna()
        
        return df_sorted.reset_index(drop=True)
    
    def remove_duplicates_multiple_sorts(self, df, columns):
        """複数の条件でソートしながら重複削除"""
        sort_columns = ["間取り", "向き", "賃料(円)"]
        
        for sort_col in sort_columns:
            if sort_col in df.columns:
                df = df.sort_values(by=sort_col, ascending=True)
                df = df.sort_values(by="物件名", ascending=True)
                df = df.loc[~df[columns].eq(df[columns].shift(-1)).all(axis=1)]
        
        return df.reset_index(drop=True)
    
    def categorize_floor_plans(self, df):
        """間取り分類"""
        df["間取り分類"] = df["間取り"].apply(
            lambda x: x if x in self.config.CATEGORIES else "その他"
        )
        
        cat1 = df.groupby("間取り分類").agg(
            件数=("間取り分類", "count"),
            平均賃料=("賃料(円)", "mean"),
            平均専有面積=("専有面積(㎡)", "mean")
        ).reset_index()
        
        cat1[["平均賃料", "平均専有面積"]] = cat1[["平均賃料", "平均専有面積"]].round(1)
        
        return cat1