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

from src.utils.driver import WebDriver
from src.utils.file_handler import FileHandler
from src.utils.parser import DataParser
from src.models.company import Company

class MOPSCrawler:
    def __init__(self):
        """初始化爬蟲類"""
        self.url = "https://mopsov.twse.com.tw/mops/web/index"
        self.driver = WebDriver.create_driver()
        self.wait = WebDriver.create_wait(self.driver)
        self.short_wait = WebDriver.create_wait(self.driver, 3)
        self.companies = FileHandler.read_company_list()
        self.results = []

    def retry_on_connection_error(self, func, max_retries=3):
        """帶重試機制的函數執行器"""
        for attempt in range(max_retries):
            try:
                return func()
            except Exception as e:
                if attempt == max_retries - 1:
                    raise e
                print(f"連線出錯，等待後重試... (第 {attempt + 1} 次)")
                time.sleep(5 * (attempt + 1))
                
                try:
                    self.driver.refresh()
                except:
                    pass
                    
                try:
                    if self.driver.current_url == "about:blank" or "autoAction" in self.driver.current_url:
                        print("檢測到主頁面錯誤，嘗試重新訪問...")
                        self.driver.get(self.url)
                        time.sleep(3)
                except:
                    pass

    def search_company(self, company):
        """搜尋公司資訊"""
        def search_action():
            main_window = None
            new_window = None
            
            try:
                # 檢查公司代碼
                if not company.code or company.code == '查無資訊':
                    logger.warning(f"公司 {company.name} 無代碼資訊")
                    return None

                print("\n" + "="*50)
                print(f"開始處理公司：{company.name}")
                print(f"公司代碼：{company.code}")
                print("="*50)

                # 搜尋公司
                search_box = self.wait.until(
                    EC.presence_of_element_located((By.ID, "keyword"))
                )
                search_box.clear()
                search_box.send_keys(company.code)
                time.sleep(1)

                # 點擊搜尋
                search_button = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "rulesubmit"))
                )
                search_button.click()
                time.sleep(3)

                # 獲取公司資訊
                main_window = self.driver.current_window_handle
                company_info = DataParser.parse_company_info(self.driver)
                if company_info:
                    company.industry = company_info.get('產業類別')
                    print(f"\n公司基本資訊：")
                    print(f"公司名稱：{company_info['公司名稱']}")
                    print(f"公司代號：{company_info['公司代號']}")
                    print(f"產業類別：{company.industry}")

                # 點擊財務報表
                self.wait.until(
                    EC.element_to_be_clickable((By.ID, "button11"))
                )

                # 提交表單
                js_code = """
                var form = document.fm1;
                form.step.value = '1';
                form.co_id.value = arguments[0];
                form.year.value = '113';
                form.season.value = '04';
                form.action = '/mops/web/ajax_t164sb04';
                form.target = '_blank';
                form.submit();
                """
                
                current_handles = set(self.driver.window_handles)
                self.driver.execute_script(js_code, company.code)
                time.sleep(2)

                # 處理新視窗
                new_window = self._wait_for_new_window(current_handles)
                if not new_window:
                    print("無法找到包含數據的視窗")
                    return None

                # 解析財務數據
                tables = self.driver.find_elements(By.CLASS_NAME, "hasBorder")
                if tables:
                    financial_data = DataParser.parse_financial_data(tables[0])
                    if financial_data:
                        company.annual_revenue = financial_data.annual_revenue
                        company.gross_profit = financial_data.gross_profit
                        company.gross_margin = financial_data.gross_margin
                        company.profit_before_tax = financial_data.profit_before_tax
                        company.profit_after_tax = financial_data.profit_after_tax
                        
                        print(f"\n財務數據：")
                        print(f"年營收：{company.annual_revenue}")
                        print(f"毛利額：{company.gross_profit}")
                        print(f"毛利率：{company.gross_margin}")
                        print(f"稅前淨利：{company.profit_before_tax}")
                        print(f"稅後淨利：{company.profit_after_tax}")
                        print("-" * 50)
                    else:
                        print("警告：未找到任何財務數據")
                else:
                    print("警告：未找到財務數據表格")

            except Exception as e:
                print(f"處理公司 {company.name} 時出錯: {str(e)}")
                return None
            
            finally:
                self._cleanup_windows(main_window)
                
            return company

        try:
            return self.retry_on_connection_error(search_action)
        except Exception as e:
            logger.error(f"搜尋公司 {company.name} ({company.code}) 時出錯: {str(e)}")
            return None

    def _wait_for_new_window(self, current_handles, max_wait=10):
        """等待新視窗開啟"""
        for _ in range(max_wait):
            try:
                new_handles = set(self.driver.window_handles) - current_handles
                if new_handles:
                    for handle in new_handles:
                        self.driver.switch_to.window(handle)
                        if "ajax_t164sb04" in self.driver.current_url:
                            return handle
                        self.driver.close()
            except Exception as e:
                print(f"切換視窗時出錯: {str(e)}")
            time.sleep(1)
        return None

    def _cleanup_windows(self, main_window):
        """清理視窗"""
        try:
            all_handles = self.driver.window_handles
            if main_window and main_window in all_handles:
                for handle in all_handles:
                    if handle != main_window:
                        self.driver.switch_to.window(handle)
                        self.driver.close()
                self.driver.switch_to.window(main_window)
            else:
                first_window = all_handles[0]
                for handle in all_handles[1:]:
                    self.driver.switch_to.window(handle)
                    self.driver.close()
                self.driver.switch_to.window(first_window)
        except Exception as e:
            print(f"清理視窗時出錯: {str(e)}")

    def crawl_all_companies(self):
        """爬取所有公司資訊"""
        try:
            def init_page():
                self.driver.get(self.url)
                self.wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
                self.wait.until(EC.presence_of_element_located((By.ID, "keyword")))
                time.sleep(2)
            
            self.retry_on_connection_error(init_page)
            
            for company in self.companies:
                retry_count = 0
                max_retries = 3
                success = False
                
                while not success and retry_count < max_retries:
                    try:
                        logger.info(f"正在爬取公司: {company.name} ({company.code or '無代碼'})")
                        
                        if "index" not in self.driver.current_url:
                            self.driver.get(self.url)
                            time.sleep(3)
                        
                        result = self.search_company(company)
                        if result:
                            self.results.append(result)
                            success = True
                        else:
                            raise Exception("未能獲取有效數據")
                            
                    except Exception as e:
                        retry_count += 1
                        if retry_count < max_retries:
                            print(f"處理公司 {company.name} 時出錯，等待後重試... (第 {retry_count} 次)")
                            time.sleep(5 * retry_count)
                            try:
                                self.driver.get(self.url)
                                time.sleep(3)
                            except:
                                pass
                        else:
                            logger.error(f"處理公司 {company.name} 失敗，已達到最大重試次數")
                
                time.sleep(3)
                
            FileHandler.save_results(self.results)
            
        except Exception as e:
            logger.error(f"爬取過程中出錯: {str(e)}")
        finally:
            self.driver.quit()

if __name__ == "__main__":
    crawler = MOPSCrawler()
    crawler.crawl_all_companies() 