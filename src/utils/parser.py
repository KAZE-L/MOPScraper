from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from src.models.company import Company
import time

class DataParser:
    """資料解析工具類"""
    @staticmethod
    def parse_company_info(driver):
        """解析公司基本資訊"""
        try:
            company_info = {}
            
            # 等待頁面載入
            time.sleep(2)
            
            # 尋找包含公司資訊的 div 元素
            try:
                # 使用 XPath 找到所有包含公司資訊的 div
                divs = driver.find_elements(By.XPATH, "//div[contains(@style, 'font-weight:bold')]")
                for div in divs:
                    text = div.text.strip()
                    if "公司名稱：" in text:
                        company_info['公司名稱'] = text.replace("公司名稱：", "").strip()
                    elif "公司代號：" in text:
                        company_info['公司代號'] = text.replace("公司代號：", "").strip()
                    elif "產業類別：" in text:
                        company_info['產業類別'] = text.replace("產業類別：", "").strip()
            except Exception as e:
                print(f"解析公司資訊 div 時出錯: {str(e)}")

            # 移除可能的空白字符
            if '公司名稱' in company_info:
                company_info['公司名稱'] = company_info['公司名稱'].strip()
            if '公司代號' in company_info:
                company_info['公司代號'] = company_info['公司代號'].strip()
            if '產業類別' in company_info:
                company_info['產業類別'] = company_info['產業類別'].strip()

            # 設置預設值
            company_info.setdefault('公司名稱', '無資料')
            company_info.setdefault('公司代號', '無資料')
            company_info.setdefault('產業類別', '無資料')

            print(f"解析到的公司資訊: {company_info}")  # 調試輸出
            return company_info

        except Exception as e:
            print(f"解析公司資訊時出錯: {str(e)}")
            return None

    @staticmethod
    def parse_financial_data(table):
        """解析財務數據"""
        try:
            company = Company()
            data_found = False
            
            rows = table.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                try:
                    cols = row.find_elements(By.TAG_NAME, "td")
                    if len(cols) >= 2:
                        title = cols[0].text.strip()
                        value = cols[1].text.strip().replace(" ", "")  # 移除數字中的空格
                        percentage = cols[2].text.strip() if len(cols) >= 3 else ""
                        
                        if "營業收入合計" in title:
                            company.annual_revenue = value
                            data_found = True
                        elif "營業毛利（毛損）淨額" in title:
                            company.gross_profit = value
                            company.gross_margin = percentage
                            data_found = True
                        elif "稅前淨利（淨損）" in title:
                            company.profit_before_tax = value
                            data_found = True
                        elif "本期淨利（淨損）" in title:
                            company.profit_after_tax = value
                            data_found = True
                except Exception as e:
                    print(f"解析行數據時出錯: {str(e)}")
                    continue
            
            return company if data_found else None
        except Exception as e:
            print(f"解析財務數據時出錯: {str(e)}")
            return None 