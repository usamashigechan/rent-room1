# config.py
import os
from datetime import datetime

class Config:
    FOLDER_PATH = r"C:\\1111accommodation"
    TIMEOUT = 10
    DELAY = 0.5
    MAX_ITEMS = 100
    CATEGORIES = ["ワンルーム", "1K", "1DK", "1LDK", "2K", "2DK", "2LDK", "3K", "3DK", "3LDK"]
    
    @staticmethod
    def get_timestamp():
        return datetime.now().strftime("%y%m%d%H%M")
    
    @staticmethod
    def get_datestamp():
        return datetime.now().strftime("%y%m%d")