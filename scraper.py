# scraper.py
import requests
import time
import numpy as np
import pandas as pd
from bs4 import BeautifulSoup
from config import Config

class PropertyScraper:
    def __init__(self):
        self.config = Config()
    
    def scrape_page(self, url):
        """単一ページのスクレイピング"""
        try:
            response = requests.get(url, timeout=self.config.TIMEOUT)
            if response.status_code != 200:
                return None
            return BeautifulSoup(response.text, "html.parser")
        except requests.exceptions.RequestException as e:
            print(f"ページ取得エラー: {e}")
            return None
    
    def extract_titles_and_links(self, soup):
        """物件名とURLの抽出"""
        titles = [title.text.strip() for title in soup.find_all("h2", class_="property_inner-title")]
        links = [a["href"] for a in soup.find_all("a", href=True) if "/chintai/bc" in a["href"]]
        full_links = ["https://suumo.jp" + link for link in links[:self.config.MAX_ITEMS]]
        return titles, full_links
    
    def extract_prices(self, soup):
        """賃料の抽出と変換"""
        prices = [title.text.strip() for title in soup.find_all("div", class_="detailbox-property-point")]
        
        def convert_price(price):
            try:
                return int(float(price.replace('万円', '')) * 10000)
            except ValueError:
                return np.nan
        
        return [convert_price(price) for price in prices]
    
    def extract_walk_times(self, soup):
        """徒歩時間の抽出"""
        walk_times = []
        detail_notes = soup.find_all("div", class_="font-weight:bold") + soup.find_all("div", style="font-weight:bold")
        
        for note in detail_notes:
            text = note.text.strip()
            try:
                if "歩" in text and "分" in text and "バス" not in text and "車" not in text:
                    walk_time_str = text.split("歩")[1].split("分")[0].strip()
                    walk_time = int(walk_time_str)
                    walk_times.append(walk_time)
                else:
                    walk_times.append(None)
            except (ValueError, IndexError):
                walk_times.append(None)
        
        return walk_times
    
    def extract_property_details(self, soup):
        """詳細情報の抽出"""
        properties = []
        for row in soup.find_all("tr")[:self.config.MAX_ITEMS]:
            try:
                property_data = {
                    "管理費": row.find("td", class_="detailbox-property-col detailbox-property--col1").find_all("div")[1].text.strip(),
                    "敷金": row.find("td", class_="detailbox-property-col detailbox-property--col2").find_all("div")[0].text.strip(),
                    "礼金": row.find("td", class_="detailbox-property-col detailbox-property--col2").find_all("div")[1].text.strip(),
                    "間取り": row.find("td", class_="detailbox-property-col detailbox-property--col3").find_all("div")[0].text.strip(),
                    "専有面積(㎡)": row.find("td", class_="detailbox-property-col detailbox-property--col3").find_all("div")[1].text.strip(),
                    "向き": row.find("td", class_="detailbox-property-col detailbox-property--col3").find_all("div")[2].text.strip(),
                    "物件種別": row.find_all("td", class_="detailbox-property-col detailbox-property--col3")[1].find_all("div")[0].text.strip(),
                    "築年数(年)": row.find_all("td", class_="detailbox-property-col detailbox-property--col3")[1].find_all("div")[1].text.strip(),
                    "住所": row.find_all("td", class_="detailbox-property-col")[-1].text.strip()
                }
                properties.append(property_data)
            except:
                continue
        return properties
    
    def scrape_station(self, station, base_url, num_pages):
        """指定駅の全ページをスクレイピング"""
        all_dataframes = []
        
        for i in range(1, num_pages + 1):
            url = base_url + str(i)
            print(f"取得中: {url}")
            
            time.sleep(self.config.DELAY)
            
            soup = self.scrape_page(url)
            if not soup:
                continue
            
            titles, full_links = self.extract_titles_and_links(soup)
            rents = self.extract_prices(soup)
            walk_times = self.extract_walk_times(soup)
            
            min_len = min(len(titles), len(full_links), len(rents), len(walk_times))
            if min_len == 0:
                continue
            
            df1 = pd.DataFrame({
                "物件名": titles[:min_len],
                "URL": full_links[:min_len],
                "賃料(円)": rents[:min_len],
                "徒歩時間(分)": walk_times[:min_len]
            })
            
            properties = self.extract_property_details(soup)
            if properties:
                df2 = pd.DataFrame(properties)
                df2["専有面積(㎡)"] = df2["専有面積(㎡)"].str.replace("m2", "").astype(float)
                df2["築年数(年)"] = pd.to_numeric(df2["築年数(年)"].str.replace("築", "").str.replace("年", "").str.replace("新築", "0"), errors="coerce").astype("Int64")
                df2["築年数(年)"] = df2["築年数(年)"].fillna(0).astype(int)
                
                df_combined = pd.concat([df1, df2], axis=1)
                all_dataframes.append(df_combined)
        
        return all_dataframes