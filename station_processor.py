# station_processor.py
import os
from datetime import datetime
import pandas as pd
import numpy as np
import statsmodels.api as sm
from config import Config

class StationProcessor:
    def __init__(self, scraper, processor, stats_calc, visualizer, ppt_gen, file_manager):
        self.scraper = scraper
        self.processor = processor
        self.stats_calc = stats_calc
        self.visualizer = visualizer
        self.ppt_gen = ppt_gen
        self.file_manager = file_manager
        self.config = Config()
    
    def process_station(self, station, base_url, num_pages):
        """駅の完全処理"""
        print(f"\n=== {station} のスクレイピング開始 ===")
        
        # スクレイピング
        all_dataframes = self.scraper.scrape_station(station, base_url, num_pages)
        
        if not all_dataframes:
            print(f"{station}: データが1件も取得できませんでした。")
            return None
        
        # データクリーニング
        df_sorted = self.processor.clean_data(all_dataframes)
        if df_sorted is None or len(df_sorted) == 0:
            print(f"{station}: クリーニング後にデータがありません。")
            return None
        
        n = len(df_sorted)
        print(f"{station}: スクレイピング完了 - {n}件保存")
        
        # ファイル保存
        datestamp = self.config.get_datestamp()
        timestamp = self.config.get_timestamp()
        
        file_name = f"1fData_{station}_{datestamp}.csv"
        self.file_manager.save_csv(df_sorted, file_name)
        
        # 統計処理・分析
        return self.perform_analysis(station, df_sorted, n, datestamp, timestamp)
    
    def perform_analysis(self, station, df_sorted, n, datestamp, timestamp):
        """統計分析とファイル生成"""
        print(f"{station}: 統計処理・グラフ作成開始")
        
        # 基本情報作成
        df_base1 = pd.DataFrame([
            ["全データ数", "取得した現在時刻", "調査駅", "出典"],
            [n, "day"+timestamp, station, "https://suumo.jp/jj/chintai"]
        ])
        
        # 統計データ作成
        stats_df, (stats11, stats12) = self.stats_calc.create_stats_dataframe(df_sorted)
        
        # CSVファイル保存
        csv_files = {
            f"{station}_{datestamp}_base1.csv": df_base1,
            f"{station}_{datestamp}_stats1.csv": stats_df,
            f"{station}_{datestamp}_stats11.csv": stats11,
            f"{station}_{datestamp}_stats12.csv": stats12
        }
        
        for filename, df in csv_files.items():
            self.file_manager.save_csv(df, filename)
        
        # 間取り分類
        cat1 = self.processor.categorize_floor_plans(df_sorted)
        self.file_manager.save_csv(cat1, f"{station}_{datestamp}_ct1.csv")
        
        # グラフ作成
        folder_path = self.file_manager.base_folder
        
        # 分布グラフ
        tg1_path = os.path.join(folder_path, f"{station}_{datestamp}_tg1.png")
        self.visualizer.create_distribution_plots(df_sorted, station, tg1_path)
        
        # 散布図
        tg2_path = os.path.join(folder_path, f"{station}_{datestamp}_tg2.png")
        self.visualizer.create_scatter_plots(df_sorted, tg2_path)
        
        # 重回帰分析
        regression_results = self.stats_calc.perform_multiple_regression(df_sorted)
        model = regression_results["model"]
        
        # 重回帰結果のCSV保存
        df_mrl1 = pd.DataFrame([
            ["指標", "値"],
            ["補正決定係数", regression_results["adj_r_squared"]],
            ["F値", regression_results["f_stat"]],
            ["Fのp値", regression_results["f_p_value"]]
        ]).T
        
        df_mrl2 = pd.DataFrame([
            ["item", "coef(切片、傾き)","p値"],
            ["切片", regression_results["intercept_coef"],"-"],
            ["徒歩時間(分)", regression_results["coefficients"]["徒歩時間(分)"], regression_results["p_values"]["徒歩時間(分)"]],
            ["築年数(年)", regression_results["coefficients"]["築年数(年)"], regression_results["p_values"]["築年数(年)"]],
            ["専有面積(㎡)", regression_results["coefficients"]["専有面積(㎡)"], regression_results["p_values"]["専有面積(㎡)"]]
        ])
        
        self.file_manager.save_csv(df_mrl1, f"{station}_{datestamp}_mrl1.csv")
        self.file_manager.save_csv(df_mrl2, f"{station}_{datestamp}_mrl2.csv")
        
        # VIF計算
        X = df_sorted[['徒歩時間(分)', '築年数(年)', '専有面積(㎡)']]
        X_with_const = sm.add_constant(X)
        df_vif1 = self.stats_calc.calculate_vif(X_with_const)
        self.file_manager.save_csv(df_vif1, f"{station}_{datestamp}_vif1.csv")
        
        # 予測値計算とプロット
        df_plot, std_residuals = self.stats_calc.create_predictions(df_sorted, model)
        
        mlrap1_path = os.path.join(folder_path, f"{station}_{datestamp}_mlrap1.png")
        self.visualizer.create_prediction_plot(df_plot, model, std_residuals, mlrap1_path)
        
        # 面積別予測
        coefficients = regression_results["coefficients"]
        intercept_coef = regression_results["intercept_coef"]
        gap_pred = std_residuals * 1.96
        
        predictions = []
        for area in [25, 50, 75, 100]:
            pred = round(intercept_coef + coefficients["専有面積(㎡)"]*area + 
                        coefficients["徒歩時間(分)"]*10 + coefficients["築年数(年)"]*10, 1)
            predictions.append([f"{area}m²", pred, round(pred - gap_pred, 1), round(pred + gap_pred, 1)])
        
        df_comp1 = pd.DataFrame(predictions, columns=["専有面積", "予測値", "予測下限", "予測上限"])
        self.file_manager.save_csv(df_comp1, f"{station}_{datestamp}_comp1.csv")
        
        # PowerPoint作成
        ppt = self.ppt_gen.create_presentation(station, n, timestamp, folder_path, 
                                             df_base1, cat1, stats11, stats12, 
                                             df_mrl1, df_mrl2, df_vif1, df_comp1)
        
        file_name_ppt = f"1e_{station}_{timestamp}_ptt1.pptx"
        file_path_ppt = os.path.join(folder_path, file_name_ppt)
        ppt.save(file_path_ppt)
        print(f"PowerPointファイルを保存しました: {file_path_ppt}")
        
        print(f"{station}: 統計処理・グラフ作成・PowerPoint作成完了")
        
        return {'station': station, 'data': df_sorted, 'count': n}