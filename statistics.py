# statistics.py
import numpy as np
import pandas as pd
import scipy.stats as stats
from scipy.stats import linregress
import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor

class StatisticsCalculator:
    def calculate_basic_stats(self, series, round_digits=2):
        """基礎統計量の計算"""
        return {
            "平均": round(series.mean(), round_digits),
            "中央値": round(series.median(), round_digits),
            "不変標準偏差": round(series.std(ddof=1), 4),
            "標準誤差": round(series.std(ddof=1) / np.sqrt(len(series)), 4),
            "最小値": series.min(),
            "最大値": series.max(),
            "第一四分位": round(series.quantile(0.25), round_digits),
            "第三四分位": round(series.quantile(0.75), round_digits),
            "尖度": round(series.kurt(), 2),
            "歪度": round(series.skew(), 2)
        }
    
    def create_stats_dataframe(self, df):
        """統計データフレームの作成"""
        columns = ["賃料(円)", "徒歩時間(分)", "専有面積(㎡)", "築年数(年)"]
        stats_data = []
        
        for col in columns:
            if col in df.columns:
                df[col] = df[col].astype(float)
                stats = self.calculate_basic_stats(df[col])
                stats["項目"] = col
                stats_data.append(stats)
        
        stats_df = pd.DataFrame(stats_data)
        stats_df = stats_df[["項目"] + [col for col in stats_df.columns if col != "項目"]]
        
        return stats_df.T, self.split_stats_dataframe(stats_df)
    
    def split_stats_dataframe(self, stats_df):
        """統計データフレームの分割"""
        stats11_cols = ["項目", "平均", "中央値", "不変標準偏差", "標準誤差"]
        stats12_cols = ["項目", "最小値", "最大値", "第一四分位", "第三四分位", "尖度", "歪度"]
        
        stats11 = stats_df[stats11_cols].T
        stats12 = stats_df[stats12_cols].T
        
        return stats11, stats12
    
    def perform_multiple_regression(self, df):
        """重回帰分析の実行"""
        X = df[['徒歩時間(分)', '築年数(年)', '専有面積(㎡)']]
        y = df['賃料(円)']
        
        X = sm.add_constant(X)
        model = sm.OLS(y, X).fit()
        
        print(model.summary())
        
        results = {
            "model": model,
            "adj_r_squared": model.rsquared_adj,
            "f_stat": model.fvalue,
            "f_p_value": model.f_pvalue,
            "intercept_coef": model.params["const"],
            "coefficients": model.params.drop("const"),
            "p_values": model.pvalues.drop("const")
        }
        
        return results
    
    def calculate_vif(self, X):
        """VIFの計算"""
        vif_data = []
        feature_names = ["徒歩時間(分)", "築年数(年)", "専有面積(㎡)"]
        
        for i in range(1, X.shape[1]):  # 定数項を除く
            vif_value = variance_inflation_factor(X.values, i)
            vif_data.append([feature_names[i-1], vif_value])
        
        return pd.DataFrame(vif_data, columns=["item", "VIF"])
    
    def create_predictions(self, df, model):
        """予測値の計算"""
        df_plot = df.copy().drop_duplicates().reset_index(drop=True)
        
        X_pred = sm.add_constant(df_plot[['徒歩時間(分)', '築年数(年)', '専有面積(㎡)']])
        df_plot['predicted_rent'] = model.predict(X_pred)
        
        residuals = df_plot['賃料(円)'] - df_plot['predicted_rent']
        std_residuals = np.std(residuals)
        
        df_plot['upper_bound'] = df_plot['predicted_rent'] + (std_residuals * 1.96)
        df_plot['lower_bound'] = df_plot['predicted_rent'] - (std_residuals * 1.96)
        
        return df_plot, std_residuals