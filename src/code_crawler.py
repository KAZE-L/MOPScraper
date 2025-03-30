from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from loguru import logger
import pandas as pd
import time
import os

class CompanyCodeCrawler:
    def __init__(self):
        self.url = "https://mops.twse.com.tw/mops/#/web/home"
        self.setup_driver()
        self.companies = self.read_company_list()
        self.results = []

    def setup_driver(self):
        """設置Chrome瀏覽器驅動"""
        chrome_options = Options()
        chrome_options.add_argument('--headless')  # 無頭模式
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.wait = WebDriverWait(self.driver, 10)

    def read_company_list(self):
        """從Excel檔案讀取公司列表"""
        try:
            df = pd.read_excel('company_list.xlsx')
            companies = df['公司名稱'].tolist()
            logger.info(f"成功讀取了 {len(companies)} 家公司的名稱")
            return companies
        except FileNotFoundError:
            logger.error("找不到company_list.xlsx檔案，將創建新檔案")
            # 如果檔案不存在，創建一個空的DataFrame並保存
            df = pd.DataFrame({'公司名稱': []})
            df.to_excel('company_list.xlsx', index=False)
            return []
        except Exception as e:
            logger.error(f"讀取公司列表時出錯: {str(e)}")
            return []

    def search_company_code(self, company_name):
        """搜尋公司代碼"""
        try:
            # 等待搜尋框可用並輸入公司名稱
            search_box = self.wait.until(
                EC.presence_of_element_located((By.ID, "searchInfo"))
            )
            search_box.clear()
            search_box.send_keys(company_name)
            time.sleep(2)  # 等待搜尋結果出現

            # 查找搜尋結果
            search_results = self.wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "searchBlock"))
            )
            
            # 獲取包含公司代碼的連結
            company_link = search_results.find_element(By.TAG_NAME, "a")
            link_text = company_link.text
            
            # 提取公司代碼（格式：代碼 公司名稱）
            company_code = link_text.split()[0]
            
            logger.info(f"找到公司 {company_name} 的代碼: {company_code}")
            return company_code

        except Exception as e:
            logger.error(f"搜尋公司 {company_name} 的代碼時出錯: {str(e)}")
            return "查無資訊"

    def crawl_all_companies(self):
        """爬取所有公司代碼"""
        try:
            self.driver.get(self.url)
            time.sleep(2)  # 等待頁面完全載入
            
            results = []
            for company_name in self.companies:
                company_code = self.search_company_code(company_name)
                results.append({
                    '公司名稱': company_name,
                    '公司代號': company_code
                })
                time.sleep(1)  # 避免請求過於頻繁
            
            # 保存結果
            df = pd.DataFrame(results)
            df.to_excel('company_list.xlsx', index=False)
            logger.info("已將公司代碼保存到company_list.xlsx")
            
        except Exception as e:
            logger.error(f"爬取過程中出錯: {str(e)}")
        finally:
            self.driver.quit()

if __name__ == "__main__":
    # 確保data目錄存在
    os.makedirs('data', exist_ok=True)
    
    # 開始爬取
    crawler = CompanyCodeCrawler()
    crawler.crawl_all_companies() 