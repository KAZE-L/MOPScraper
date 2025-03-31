import os
import pandas as pd
from loguru import logger
from src.models.company import Company

class FileHandler:
    """檔案處理工具類"""
    @staticmethod
    def read_company_list(file_path='data/input/company_list.xlsx'):
        """從Excel檔案讀取公司列表"""
        try:
            df = pd.read_excel(file_path)
            companies = df.to_dict('records')
            logger.info(f"成功讀取了 {len(companies)} 家公司的資訊")
            return [Company.from_dict(company) for company in companies]
        except FileNotFoundError:
            logger.error(f"找不到 {file_path} 檔案")
            return []
        except Exception as e:
            logger.error(f"讀取公司列表時出錯: {str(e)}")
            return []

    @staticmethod
    def save_results(results, output_path='data/output/results.xlsx'):
        """保存爬取結果"""
        try:
            if not results:
                logger.warning("沒有數據可保存")
                return

            # 將Company物件轉換為字典
            data = [company.to_dict() for company in results]
            
            # 創建DataFrame
            df = pd.DataFrame(data)
            
            # 調整列順序
            columns_order = ['公司名稱', '公司代號', '產業類別', '年營收', 
                           '毛利額', '毛利率', '稅前淨利', '稅後淨利']
            df = df[columns_order]
            
            # 確保輸出目錄存在
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # 保存為Excel檔案
            df.to_excel(output_path, index=False)
            logger.success(f"結果已保存至: {output_path}")
            
        except Exception as e:
            logger.error(f"保存結果時出錯: {str(e)}") 