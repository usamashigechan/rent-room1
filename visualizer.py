# visualizer.py
import matplotlib.pyplot as plt
import numpy as np
from scipy.stats import linregress

class Visualizer:
    def __init__(self):
        plt.rcParams['font.family'] = 'MS Gothic'
    
    def create_distribution_plots(self, df, station, output_path):
        """分布グラフの作成"""
        columns = ["賃料(円)", "徒歩時間(分)", "専有面積(㎡)", "築年数(年)"]
        colors = ["skyblue", "lightgreen", "lightgreen", "lightgreen"]
        
        # 統合グラフ
        fig, axes = plt.subplots(4, 2, figsize=(12, 12))
        
        for i, (col, color) in enumerate(zip(columns, colors)):
            if col in df.columns:
                # ヒストグラム
                axes[i, 0].hist(df[col], bins=30, edgecolor='black')
                axes[i, 0].set_title(f"{col}のヒストグラム")
                axes[i, 0].set_xlabel(col)
                axes[i, 0].set_ylabel("度数")
                
                # 箱ひげ図
                axes[i, 1].boxplot(df[col], patch_artist=True, boxprops=dict(facecolor=color))
                axes[i, 1].set_title(f"{col}の箱ひげ図")
                axes[i, 1].set_xlabel(station)
                axes[i, 1].set_ylabel(col)
        
        plt.tight_layout()
        plt.savefig(output_path)
        plt.close()
        
        # 個別グラフの作成
        self.create_individual_plots(df, station, columns, colors, output_path)
    
    def create_individual_plots(self, df, station, columns, colors, base_path):
        """個別グラフの作成"""
        suffixes = ["r1", "w1", "s1", "a1"]
        
        for col, color, suffix in zip(columns, colors, suffixes):
            if col in df.columns:
                fig, axes = plt.subplots(1, 2, figsize=(12, 6))
                
                axes[0].hist(df[col], bins=30, edgecolor='black')
                axes[0].set_title(f"{col}のヒストグラム")
                axes[0].set_xlabel(col)
                axes[0].set_ylabel("度数")
                
                axes[1].boxplot(df[col], patch_artist=True, boxprops=dict(facecolor=color))
                axes[1].set_title(f"{col}の箱ひげ図")
                axes[1].set_xlabel(station)
                axes[1].set_ylabel(col)
                
                plt.tight_layout()
                file_path = base_path.replace("_tg1.png", f"_g{suffix}.png")
                plt.savefig(file_path)
                plt.close()
    
    def create_scatter_plots(self, df, output_path):
        """散布図の作成"""
        x_data = [df["徒歩時間(分)"], df["専有面積(㎡)"], df["築年数(年)"]]
        y_data = [df["賃料(円)"], df["賃料(円)"], df["賃料(円)"]]
        titles = ["賃料(円) vs 徒歩時間(分)", "専有面積(㎡) vs 賃料(円)", "賃料(円) vs 築年数(年)"]
        x_labels = ["徒歩時間(分)", "専有面積(㎡)", "築年数(年)"]
        y_labels = ["賃料(円)", "賃料(円)", "賃料(円)"]
        
        # 統合散布図
        fig, axes = plt.subplots(3, 1, figsize=(6, 18))
        
        for i in range(3):
            slope, intercept, r_value, p_value, std_err = linregress(x_data[i], y_data[i])
            line_eq = f"y = {slope:.2f}x + {intercept:.2f}"
            
            axes[i].scatter(x_data[i], y_data[i], alpha=0.5, color="blue", label="データ")
            axes[i].plot(x_data[i], slope*x_data[i] + intercept, color="red", label=f"近似直線: {line_eq}")
            axes[i].text(min(x_data[i]), max(y_data[i]), f"R² = {r_value**2:.2f}\n p値 = {p_value:.4f}", fontsize=10, color="black")
            axes[i].set_title(titles[i])
            axes[i].set_xlabel(x_labels[i])
            axes[i].set_ylabel(y_labels[i])
            axes[i].legend(loc="lower right")
        
        plt.tight_layout()
        plt.savefig(output_path)
        plt.close()
        
        # 個別散布図の作成
        self.create_individual_scatter_plots(x_data, y_data, titles, x_labels, y_labels, output_path)
    
    def create_individual_scatter_plots(self, x_data, y_data, titles, x_labels, y_labels, base_path):
        """個別散布図の作成"""
        for i in range(3):
            plt.figure(figsize=(6, 4))
            
            slope, intercept, r_value, p_value, std_err = linregress(x_data[i], y_data[i])
            line_eq = f"y = {slope:.2f}x + {intercept:.2f}"
            
            plt.scatter(x_data[i], y_data[i], alpha=0.5, color="blue", label="データ")
            plt.plot(x_data[i], slope*x_data[i] + intercept, color="red", label=f"近似直線: {line_eq}")
            plt.text(min(x_data[i]), max(y_data[i]), f"R² = {r_value**2:.2f}\n p値 = {p_value:.4f}", fontsize=10, color="black")
            
            plt.title(titles[i])
            plt.xlabel(x_labels[i])
            plt.ylabel(y_labels[i])
            plt.legend(loc="lower right")
            
            file_path = base_path.replace("_tg2.png", f"_tgscat{i+1}.png")
            plt.savefig(file_path)
            plt.close()
    
    def create_prediction_plot(self, df_plot, model, std_residuals, output_path):
        """予測vs実測のプロット作成"""
        n_samples = len(df_plot)
        slope, intercept = np.polyfit(df_plot['賃料(円)'], df_plot['predicted_rent'], 1)
        line_eq = f"y = {slope:.2f}x + {intercept:.2f}"
        gap_pred = std_residuals * 1.96
        
        plt.figure(figsize=(12, 8))
        
        plt.scatter(df_plot['賃料(円)'], df_plot['predicted_rent'], 
                   color="blue", alpha=0.6, label="実測値", s=30)
        
        x_smooth = np.linspace(df_plot['賃料(円)'].min(), df_plot['賃料(円)'].max(), 100)
        y_smooth = slope * x_smooth + intercept
        
        plt.plot(x_smooth, y_smooth, "r-", lw=2, label="回帰直線")
        
        upper_smooth = y_smooth + gap_pred
        lower_smooth = y_smooth - gap_pred
        
        plt.plot(x_smooth, upper_smooth, "k--", lw=1.5, alpha=0.8, label="予測区間上限")
        plt.plot(x_smooth, lower_smooth, "k--", lw=1.5, alpha=0.8, label="予測区間下限")
        
        plt.fill_between(x_smooth, lower_smooth, upper_smooth, 
                         color="orange", alpha=0.2, label="予測区間")
        
        confidence_interval = std_residuals * 1.96 / np.sqrt(n_samples)
        upper_conf = y_smooth + confidence_interval
        lower_conf = y_smooth - confidence_interval
        
        plt.fill_between(x_smooth, lower_conf, upper_conf, 
                         color="blue", alpha=0.3, label="95% 信頼区間")
        
        plt.xlabel("実際の賃料 (円)", fontsize=12)
        plt.ylabel("予測賃料 (円)", fontsize=12)
        plt.title("実際の賃料 vs 予測賃料（信頼区間・予測区間付き）", fontsize=14)
        plt.legend(loc='upper left')
        plt.grid(True, alpha=0.3)
        
        r_squared = model.rsquared
        p_values_model = model.pvalues
        
        plt.text(0.98, 0.02, 
                 f"近似式: {line_eq}\nR² = {r_squared:.3f}\np値 = {p_values_model[1]:.3f}\nn = {n_samples}",
                 fontsize=11, verticalalignment="bottom", horizontalalignment="right",
                 transform=plt.gca().transAxes,
                 bbox=dict(facecolor="white", alpha=0.8, edgecolor="gray"))
        
        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        plt.close()