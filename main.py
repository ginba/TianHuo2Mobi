#!/usr/bin/env python3
"""
自動上傳txt和epub文件到 https://ebook.cdict.info/mobi/ 進行mobi轉換
自動勾選"強制繁體"選項，並處理廣告彈窗
修正版：等待時間調整到正確的位置
"""

import os
import time
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException


class EbookConverter:
    def __init__(self, source_folder, output_folder=None):
        """
        初始化轉換器
        
        Args:
            source_folder: 包含txt和epub文件的源文件夾路徑
            output_folder: 下載mobi文件的目標文件夾路徑(默認為源文件夾下的'mobi_output')
        """
        self.source_folder = Path(source_folder)
        self.output_folder = Path(output_folder) if output_folder else self.source_folder / 'mobi_output'
        self.output_folder.mkdir(exist_ok=True)
        
        # 配置Chrome選項
        chrome_options = Options()
        chrome_options.add_experimental_option('prefs', {
            'download.default_directory': str(self.output_folder.absolute()),
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': True
        })
        
        # 添加廣告攔截選項
        chrome_options.add_argument('--disable-popup-blocking')
        chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # 可選：使用無頭模式
        chrome_options.add_argument('--headless')
        
        self.driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )
        
        # 最大化窗口
        self.driver.maximize_window()
        
        # ========================================
        # 等待時間設定
        # ========================================
        self.wait_timeout = 120        # WebDriverWait的超時時間(秒)
        self.page_load_wait = 5        # 頁面加載後的等待時間(秒)
        self.after_file_select = 2     # 選擇文件後的短暫等待(秒)
        self.after_click_wait = 2      # 點擊按鈕後的等待時間(秒)
        self.after_submit_wait = 300   # ⭐ 提交表單(開始上傳)後的等待時間(秒) - 5分鐘
        self.conversion_wait = 600     # ⭐ 等待轉換完成的最長時間(秒) - 10分鐘
        self.download_wait = 10        # 點擊下載後等待下載完成的時間(秒)
        
        self.wait = WebDriverWait(self.driver, self.wait_timeout)
    
    def close_ads(self):
        """嘗試關閉所有可能的廣告彈窗"""
        print("    檢查並關閉廣告...")
        
        close_selectors = [
            "button.close", "a.close", "div.close",
            "[class*='close']", "[class*='Close']",
            "[id*='close']", "[id*='Close']",
            "button[aria-label='Close']",
            "button[aria-label='close']",
            ".modal-close", ".popup-close",
            "[onclick*='close']",
        ]
        
        closed_count = 0
        for selector in close_selectors:
            try:
                elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                for element in elements:
                    try:
                        if element.is_displayed():
                            element.click()
                            closed_count += 1
                            time.sleep(0.5)
                    except:
                        pass
            except:
                pass
        
        if closed_count > 0:
            print(f"    ✓ 關閉了 {closed_count} 個廣告/彈窗")
        
        time.sleep(0.5)
    
    def safe_click(self, element, description="元素"):
        """安全地點擊元素"""
        try:
            element.click()
            print(f"    ✓ {description}已點擊")
            return True
        except ElementClickInterceptedException:
            print(f"    ⚠ {description}被遮擋，嘗試使用JavaScript點擊...")
            try:
                self.driver.execute_script("arguments[0].click();", element)
                print(f"    ✓ {description}已點擊(使用JavaScript)")
                return True
            except Exception as e:
                print(f"    ✗ 無法點擊{description}: {str(e)}")
                return False
    
    def safe_send_keys(self, element, text, description="輸入框"):
        """安全地輸入文本"""
        try:
            element.clear()
            element.send_keys(text)
            print(f"    ✓ {description}已填寫: {text}")
            return True
        except:
            print(f"    ⚠ {description}被遮擋，嘗試使用JavaScript輸入...")
            try:
                self.driver.execute_script("arguments[0].value = arguments[1];", element, text)
                print(f"    ✓ {description}已填寫: {text} (使用JavaScript)")
                return True
            except Exception as e:
                print(f"    ✗ 無法填寫{description}: {str(e)}")
                return False
    
    def get_files_to_convert(self):
        """獲取所有需要轉換的txt和epub文件"""
        files = []
        for ext in ['*.txt', '*.epub']:
            files.extend(self.source_folder.glob(ext))
        return sorted(files)
    
    def convert_file(self, file_path):
        """轉換單個文件"""
        print(f"\n正在處理: {file_path.name}")
        
        try:
            # 步驟1: 訪問網站
            print("  步驟1: 正在訪問網站...")
            self.driver.get('https://ebook.cdict.info/mobi/')
            time.sleep(self.page_load_wait)
            self.close_ads()
            
            # 步驟2: 找到並選擇文件
            print("  步驟2: 正在選擇文件...")
            file_input = self.wait.until(
                EC.presence_of_element_located((By.ID, "txt_file"))
            )
            file_input.send_keys(str(file_path.absolute()))
            print(f"    ✓ 文件已選擇: {file_path.name}")
            time.sleep(self.after_file_select)
            self.close_ads()
            
            # 步驟3: 填寫書名
            print("  步驟3: 填寫書名...")
            try:
                title_input = self.driver.find_element(By.ID, "title")
                self.safe_send_keys(title_input, file_path.stem, "書名")
            except Exception as e:
                print(f"    ⚠ 書名設置失敗: {str(e)}")
            
            time.sleep(1)
            self.close_ads()
            
            # 步驟4: 選擇"強制繁體"
            print("  步驟4: 選擇'強制繁體'選項...")
            try:
                # 直接查找name="country" value="tw"的radio按鈕
                force_traditional_radio = self.driver.find_element(
                    By.CSS_SELECTOR, 
                    "input[type='radio'][name='country'][value='tw']"
                )
                
                if not force_traditional_radio.is_selected():
                    self.safe_click(force_traditional_radio, "'強制繁體'單選按鈕")
                    time.sleep(0.5)
                    
                    if force_traditional_radio.is_selected():
                        print("    ✓ '強制繁體'已成功選擇")
                    else:
                        print("    ⚠ '強制繁體'點擊後仍未選中，嘗試使用JavaScript...")
                        self.driver.execute_script("arguments[0].checked = true;", force_traditional_radio)
                        if force_traditional_radio.is_selected():
                            print("    ✓ '強制繁體'已成功選擇 (使用JavaScript)")
                        else:
                            print("    ✗ 無法選擇'強制繁體'")
                else:
                    print("    ✓ '強制繁體'已經是選中狀態")
                    
            except Exception as e:
                print(f"    ✗ 選擇'強制繁體'時出錯: {str(e)}")
                print("    提示: 可能需要手動選擇")
            
            time.sleep(1)
            self.close_ads()
            
            # 步驟5: 點擊"下一步"按鈕
            print("  步驟5: 點擊'下一步'按鈕...")
            self.close_ads()
            
            next_button = self.wait.until(
                EC.element_to_be_clickable((By.ID, "nextbutton"))
            )
            self.safe_click(next_button, "'下一步'按鈕")
            time.sleep(self.after_click_wait)
            
            # 步驟6: 等待並點擊"開始上傳"按鈕
            print("  步驟6: 等待並點擊'開始上傳'按鈕...")
            self.close_ads()
            
            submit_button = self.wait.until(
                EC.visibility_of_element_located((By.ID, "submit_button"))
            )
            time.sleep(1)
            
            self.safe_click(submit_button, "'開始上傳'按鈕")
            
            # ⭐ 重要：現在才開始等待5分鐘，讓文件上傳和處理
            print(f"\n  ⏱️  等待 {self.after_submit_wait//60} 分鐘讓文件上傳和處理...")
            print(f"      (已點擊'開始上傳'，文件正在上傳中...)")
            time.sleep(self.after_submit_wait)
            
            # 步驟7: 等待轉換完成
            print(f"\n  步驟7: 等待轉換完成(最多 {self.conversion_wait//60} 分鐘)...")
            print("          (轉換過程可能需要較長時間...)")
            
            long_wait = WebDriverWait(self.driver, self.conversion_wait)
            
            try:
                download_area = long_wait.until(
                    EC.presence_of_element_located((By.ID, "zone5"))
                )
                
                long_wait.until(
                    lambda driver: "下載 MOBI" in driver.find_element(By.ID, "zone5").text
                )
                
                print("    ✓ 轉換成功完成!")
                
                # 步驟8: 下載文件
                print("  步驟8: 正在下載文件...")
                download_button = self.driver.find_element(By.ID, "download_button")
                self.safe_click(download_button, "下載按鈕")
                
                time.sleep(self.download_wait)
                print(f"    ✓ 下載完成!")
                
                return True
                
            except TimeoutException:
                print(f"    ✗ 轉換超時 (等待了 {self.conversion_wait//60} 分鐘)")
                
                # 打印當前頁面狀態
                print("\n    當前頁面狀態:")
                for zone_id in ['zone1', 'zone2', 'zone3', 'zone4', 'zone5']:
                    try:
                        zone = self.driver.find_element(By.ID, zone_id)
                        if zone.text:
                            print(f"    {zone_id}: {zone.text[:100]}...")
                    except:
                        pass
                
                return False
            
        except Exception as e:
            print(f"  ✗ 錯誤: {file_path.name} - {str(e)}")
            import traceback
            traceback.print_exc()
            
            # 保存錯誤截圖
            try:
                screenshot_path = self.output_folder / f"error_{file_path.stem}.png"
                self.driver.save_screenshot(str(screenshot_path))
                print(f"  錯誤截圖已保存: {screenshot_path}")
            except:
                pass
            
            return False
    
    def convert_all(self):
        """轉換所有文件"""
        files = self.get_files_to_convert()
        
        if not files:
            print("沒有找到txt或epub文件")
            return
        
        print(f"\n找到 {len(files)} 個文件需要轉換")
        print(f"輸出文件夾: {self.output_folder}")
        print(f"\n⏱️  等待時間設定:")
        print(f"   • 點擊'開始上傳'後等待: {self.after_submit_wait//60} 分鐘")
        print(f"   • 轉換完成等待: {self.conversion_wait//60} 分鐘")
        print(f"   • 將自動勾選'強制繁體'選項")
        print(f"   • 將自動處理廣告彈窗")
        print("\n" + "="*60)
        
        success_count = 0
        failed_files = []
        
        for i, file_path in enumerate(files, 1):
            print(f"\n{'='*60}")
            print(f"[{i}/{len(files)}] 處理文件")
            print(f"{'='*60}")
            
            if self.convert_file(file_path):
                success_count += 1
            else:
                failed_files.append(file_path.name)
            
            if i < len(files):
                print("\n準備處理下一個文件...")
                time.sleep(3)
        
        print("\n" + "="*60)
        print(f"✓ 批量轉換完成!")
        print(f"成功: {success_count}/{len(files)}")
        if failed_files:
            print(f"\n失敗的文件:")
            for fname in failed_files:
                print(f"  • {fname}")
        print(f"\n輸出文件夾: {self.output_folder}")
        print("="*60)
    
    def close(self):
        """關閉瀏覽器"""
        print("\n按Enter鍵關閉瀏覽器...")
        input()
        if self.driver:
            self.driver.quit()


def main():
    print("="*60)
    print("MOBI 轉換器 - 修正版")
    print("="*60)
    print("\n功能說明:")
    print("  • 處理文件夾中的所有txt和epub文件")
    print("  • 自動關閉廣告彈窗")
    print("  • 自動填寫書名(使用文件名)")
    print("  • 自動勾選'強制繁體'選項")
    print("  • 點擊'開始上傳'後等待5分鐘")
    print("  • 轉換等待最多10分鐘")
    print("  • 自動下載mobi文件")
    print("="*60)
    
    #source_folder = input("\n請輸入包含txt和epub文件的文件夾路徑: ").strip()
    source_folder = "/Users/ginleikarma/Downloads/直排書預處理rar-txt"

    if not source_folder:
        print("錯誤: 請提供文件夾路徑")
        return
    
    if not os.path.exists(source_folder):
        print(f"錯誤: 文件夾不存在: {source_folder}")
        return
    
    '''output_folder = input("請輸入輸出文件夾路徑(直接按Enter使用默認路徑): ").strip()
    output_folder = output_folder if output_folder else None
    我改成下面的一行，以簡化步驟。
    '''
    output_folder = "/Users/ginleikarma/Downloads/直排書mobi"

    converter = EbookConverter(source_folder, output_folder)
    
    try:
        converter.convert_all()
    finally:
        converter.close()


if __name__ == '__main__':
    main()