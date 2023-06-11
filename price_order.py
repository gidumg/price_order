
import pandas as pd       
import openpyxl       
from selenium import webdriver       
from selenium.webdriver.common.keys import Keys       
from selenium.webdriver import ActionChains       
from bs4 import BeautifulSoup       
from tkinter import filedialog, Tk       
import time       
from selenium.webdriver.support.ui import WebDriverWait       
from selenium.webdriver.support import expected_conditions as EC       
from selenium.webdriver.common.by import By  
from selenium import webdriver       
import os       
from tqdm import tqdm       
from selenium.common.exceptions import TimeoutException       
from selenium.webdriver.chrome.options import Options       
import matplotlib.font_manager as fm  
import matplotlib.pyplot as plt  
import datetime       
from selenium.webdriver.chrome.service import Service    
from selenium.common.exceptions import UnexpectedAlertPresentException  


def naver(soup) :      
    try :      
        seller = soup.find("span", attrs={"class": "KasFrJs3SA"}).get_text()      
        uploaded_price = soup.find("span", attrs={'class' :'_1LY7DqCnwR' }).get_text()      
        platform = "스마트스토어"       
    except AttributeError :     
        seller = "확인필요"     
        uploaded_price = "확인필요"     
        platform = "스마트스토어"         
    return seller, uploaded_price, platform     
def coupang(soup) :     
    try :      
        seller = soup.find("a", attrs={"class": "prod-sale-vendor-name"}).get_text()      
        uploaded_price = soup.find("span", attrs={'class' :'total-price' }).get_text()      
        platform = "쿠팡"            
    except AttributeError :     
        seller = "확인필요"     
        uploaded_price = "확인필요"     
        platform = "쿠팡"     
    return seller, uploaded_price, platform       
       
def interpark(soup):           
    time.sleep(5)     
    try:     
        seller = soup.find("div", attrs={"class": "sellerName"}).get_text()     
        uploaded_price = soup.find("span", attrs={'class': 'priceWrap'}).get_text()     
        platform = "인터파크"     
    except AttributeError:     
        seller = "확인필요"     
        uploaded_price = "확인필요"     
        platform = "인터파크"     
    return seller, uploaded_price, platform     
def auction(soup) :     
    try :     
        seller = soup.find("a", attrs={"class": "link__seller sp_vipgroup--before sp_vipgroup--after"}).get_text()       
        uploaded_price = soup.find("strong", attrs={"class" :"price_real" }).get_text()      
        platform = "옥션/G마켓"      
    except AttributeError :     
        seller = "확인필요"     
        uploaded_price = "확인필요"     
        platform = "옥션/G마켓"     
    return seller, uploaded_price, platform    
      
def gmarket(soup) :     
    try:     
        seller = soup.find("a", attrs={"class": "link__seller sp_vipgroup--before sp_vipgroup--after"}).get_text()     
        uploaded_price = soup.find("strong", attrs={'class': 'price_real'}).get_text()     
        platform = "G마켓/옥션"     
    except AttributeError:     
        seller = "확인필요"     
        uploaded_price = "확인필요"     
        platform = "G마켓/옥션"     
    return seller, uploaded_price, platform     
def ssg(soup):     
    try:     
        seller = soup.find("a", attrs={"class": "cdtl_info_tit_link"}).get_text()     
        uploaded_price = soup.find("em", attrs={'class': 'ssg_price'}).get_text()     
        platform = "SSG"     
    except AttributeError:     
        seller = "확인필요"     
        uploaded_price = "확인필요"     
        platform = "SSG닷컴"     
    return seller, uploaded_price, platform     
def elevenst(soup):     
    try:     
        seller = soup.find("h1", attrs={"class": "c_product_store_title"}).find('a').get_text().strip()     
        uploaded_price = soup.find("span", attrs={'class': 'num value'}).get_text()     
        platform = "11번가"     
    except AttributeError:     
        seller = "확인필요"     
        uploaded_price = "확인필요"     
        platform = "11번가"     
    return seller, uploaded_price, platform    
      
def wemakeprice(soup):     
    try:     
        seller = soup.find("span", attrs={"class": "store_name"}).get_text()     
        uploaded_price = soup.find("em", attrs={'class': 'num'}).get_text()     
        platform = "위메프"     
    except AttributeError:     
        seller = "확인필요"     
        uploaded_price = "확인필요"     
        platform = "위메프"     
    return seller, uploaded_price, platform     
def lotte(soup):     
    try:     
        seller = soup.find("div", attrs={"class": "top"}).get_text()     
        uploaded_price = soup.find("span", attrs={'class': 'won'}).get_text().strip()     
        platform = "롯데ON"     
    except AttributeError:     
        seller = "확인필요"     
        uploaded_price = "확인필요"     
        platform = "롯데ON"     
    return seller, uploaded_price, platform     
def unknown(soup):     
    seller = "*확인필요*"     
    uploaded_price = ""     
    platform = ""     
    return seller, uploaded_price, platform     
def task1() :     
    base_url = 'https://search.shopping.naver.com/search/all?query='  
    keyword = input("모델명을 입력해주세요 : ")       
    comsmart_standard = int(input("지도가를 입력해주세요 : "))       
    minimum_price = int(input("최소가격을 입력해주세요 :"))     
    maximum_price = int(input("최대가격을 입력해주세요 :")) 
    options = Options()       
    options = webdriver.ChromeOptions()       
    driver = webdriver.Chrome('./mnt/c/chromedriver.exe', options=options)       
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": """ Object.defineProperty(navigator, 'webdriver', { get: () => undefined }) """})     
    url = base_url + keyword       
                   
               
    driver.get(url)       
    time.sleep(2)       
                
    search_btns = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[data-testid='SEARCH_SUB_FILTER_ORDER_LIST']")))       
            
    # Find the "낮은 가격순" button and click it       
    for btn in search_btns:       
        if btn.text.strip() == "낮은 가격순":       
            btn.click()       
            break       
                
    # 최적화 필터 제거하기       
    next_link = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a.subFilter_btn_radio__13PEL')))       
    next_link.click()       
            
    time.sleep(1)       
                
            
    min_price = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@title='최소가격 입력']")))     
    min_price.clear()     
    min_price.send_keys(minimum_price)     
            
    time.sleep(1)     
            
            
    st_date = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@title='최대가격 입력']")))      
    st_date.clear()      
    st_date.send_keys(maximum_price)      
            
    time.sleep(1)     
            
    # 가격대 검색 버튼 요소를 찾습니다.     
    search_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[@class='filter_price_srh__D6arg' and @role='button']")))      
    search_button.click()        
    time.sleep(1)         
            
            
    # 스크롤 다운       
    for c in range(0, 20):       
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight/20*" + str(c+1) + ");")       
        time.sleep(0.1)       
                    
    time.sleep(3)       
    html = driver.page_source       
    soup = BeautifulSoup(html, 'html.parser')        
    des_list = []
    link_list = []
    seller_list = []
    prc_list = []
    gap_list = []
    comment = []
    platform_list = []
    image_list = soup.select("div.product_item__MDtDF")        
            
    for item in image_list:
        for des in item.select("a.product_link__TrAac.linkAnchor"):
            des_text = des['title']
            des_list.append(des_text)
            link_text = des['href']
            link_list.append(link_text)
    for item in image_list:
        for seller in item.select_one("div.product_mall_title__Xer1m > a"):
            seller_text = seller.get_text()
            seller_list.append(seller_text)
    for item in image_list:        
        for price in item.select("span.price_num__S2p_v"):
            prc = int(price.get_text().replace("원", "").replace(",", ""))
            prc_list.append(prc)
            gap = comsmart_standard - prc
            gap_list.append(gap)
            comment_text = "안녕하세요. 컴스마트 관리부입니다. 현재 업로드 하신 제품의 당사 지도가는 " + str(comsmart_standard) + "원으로 현재 업로드하신 금액과는 당사 지도가 대비" + str(gap) + "원 차이가 있으니 판매 단가 수정을 부탁드립니다." 
            comment.append(comment_text)  
    flatform = ""
    df = pd.DataFrame({  
        "모델명" : [keyword]*len(des_list),               
        "상세정보": des_list,                
        "업체등록가": prc_list, 
        "지도가" : [comsmart_standard]*len(des_list),             
        "가격차이": gap_list,  
        "판매처" : seller_list,  
        "플랫폼" : [flatform]*len(des_list),
        "링크주소": link_list,
        "안내문구" : comment,   
        })     
    df = df[ df['가격차이'] >= 1]
    for i in range(len(df)): 
        if df.loc[i, '판매처'] == "":
            driver.get(df.loc[i, '링크주소'])
            time.sleep(8)       
            html = driver.page_source       
            soup = BeautifulSoup(html, 'lxml')                                                          
                    
            if "coupang" in html :       
                seller, uploaded_price, platform = coupang(soup)     
                    
                    
            elif '인터파크쇼핑' in html :                                      
                seller, uploaded_price, platform = interpark(soup)     
                    
                    
            elif "G마켓" in html :       
                seller, uploaded_price, platform = gmarket(soup)     
                    
                    
            elif "옥션" in html :       
                seller, uploaded_price, platform = auction(soup)     
                    
                    
            elif "ssg.com" in html :       
                seller, uploaded_price, platform = ssg(soup)     
                    
            elif "11번가" in html :        
                seller, uploaded_price, platform = elevenst(soup)     
                    
                    
            elif "위메프" in html :        
                seller, uploaded_price, platform = wemakeprice(soup)     
                    
            elif "lotte" in html :       
                seller, uploaded_price, platform = lotte(soup)                          
                    
            else :       
                seller, uploaded_price, platform = unknown(soup)  
            df.loc[i,'판매처'] = seller
            df.loc[i,'플랫폼'] = platform
    # '가격지도' 폴더 경로 생성
    dir_path = os.path.join('C:\\', '가격지도')
    # '가격지도' 폴더가 없으면 생성
    if not os.path.exists(dir_path):
        os.mkdir(dir_path)
    # 현재 시간으로 파일명 생성
    now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f'가긱지도_{now}.xlsx'
    # 파일 경로 생성
    file_path = os.path.join(dir_path, filename)
    # 파일 저장
    df.to_excel(file_path, index=False)
    print("1번 작업이 완료되었습니다.")
      
      
      
def task2() :     
    
    
    dfs = []
                
    # 엑셀 파일 읽어오기       
    file_path = filedialog.askopenfilename(title="가격지도 파일", defaultextension=".xlsx")       
    df_input = pd.read_excel(file_path, header =0)       
                            
                
    base_url = 'https://search.shopping.naver.com/search/all?query='       
                
    options = Options()       
    options = webdriver.ChromeOptions()       
    driver = webdriver.Chrome('./mnt/c/chromedriver.exe', options=options)       
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": """ Object.defineProperty(navigator, 'webdriver', { get: () => undefined }) """})       
            
            
            
                
    for i in tqdm(range(len(df_input))):       
        # 검색할 키워드를 주문번호로 대체하여 url 생성       
        keyword = str(df_input.loc[i, '주문번호'])       
        minimum_price = int(df_input.loc[i, '최소값'])       
        maximum_price = int(df_input.loc[i, '최대값'])       
                    
        #시트에서 지도가 추출하기       
        comsmart_standard = int(df_input.loc[i, '지도가'])       
                    
        url = base_url + keyword       
                    
        driver.get(url)       
        time.sleep(2)       
                    
        search_btns = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[data-testid='SEARCH_SUB_FILTER_ORDER_LIST']")))       
                    
        for btn in search_btns:       
            if btn.text.strip() == "낮은 가격순":       
                btn.click()       
                break       
                    
        #최적화 필터 제거하기           
        next_link = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a.subFilter_btn_radio__13PEL')))       
        if 'on' in next_link.get_attribute('class'):       
            next_link.click()       
                        
        time.sleep(3)       
                      
                
        min_price = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@title='최소가격 입력']")))     
        min_price.clear()     
        min_price.send_keys(minimum_price)     
            
        time.sleep(1)     
            
        st_date = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@title='최대가격 입력']")))     
        st_date.clear()      
        st_date.send_keys(maximum_price)      
            
        time.sleep(1)     
            
        search_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[@class='filter_price_srh__D6arg' and @role='button']")))        
        search_button.click()     
            
        time.sleep(1)     
                   
        for c in range(0, 20):       
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight/20*" + str(c+1) + ");")       
                time.sleep(0.1)       
                        
        time.sleep(3)       
        des_list = []
        link_list = []
        seller_list = []
        prc_list = []
        gap_list = []
        comment = []
        platform_list = []  
        
        html = driver.page_source       
        soup = BeautifulSoup(html, 'html.parser')        
        flatform = ""
        image_list = soup.select("div.product_item__MDtDF")        
                
        for item in image_list:
            for des in item.select("a.product_link__TrAac.linkAnchor"):
                des_text = des['title']
                des_list.append(des_text)
                link_text = des['href']
                link_list.append(link_text)
        for item in image_list:
            for seller in item.select_one("div.product_mall_title__Xer1m > a"):
                seller_text = seller.get_text()
                seller_list.append(seller_text)
        for item in image_list:        
            for price in item.select("span.price_num__S2p_v"):
                prc = int(price.get_text().replace("원", "").replace(",", ""))
                prc_list.append(prc)
                gap = comsmart_standard - prc
                gap_list.append(gap)
                comment_text = "안녕하세요. 컴스마트 관리부입니다. 현재 업로드 하신 제품의 당사 지도가는 " + str(comsmart_standard) + "원으로 현재 업로드하신 금액과는 당사 지도가 대비" + str(gap) + "원 차이가 있으니 판매 단가 수정을 부탁드립니다." 
                comment.append(comment_text)  
        df = pd.DataFrame({  
            "모델명" : [keyword]*len(des_list),               
            "상세정보": des_list,                
            "업체등록가": prc_list, 
            "지도가" : [comsmart_standard]*len(des_list),             
            "가격차이": gap_list,  
            "판매처" : seller_list,  
            "플랫폼" : [flatform]*len(des_list),
            "링크주소": link_list,
            "안내문구" : comment,   
            })     
        
        dfs.append(df.copy())
    df_total = pd.concat(dfs, ignore_index=True)
    df_total = df_total[ df_total['판매처'] != '쇼핑몰별 최저가' ]
    ddf_total = df_total[df_total['가격차이'] >= 1].reset_index(drop=True)
    try : 
        for i, row in ddf_total.iterrows():
            if row['판매처'] == "":
                driver.get(row['링크주소'])
                time.sleep(5)       
                html = driver.page_source       
                soup = BeautifulSoup(html, 'lxml')                                                           
                        
                if "coupang" in html :       
                    seller, uploaded_price, platform = coupang(soup)     
                        
                        
                elif '인터파크쇼핑' in html :                                 
                    seller, uploaded_price, platform = interpark(soup)     
                        
                        
                elif "G마켓" in html :       
                    seller, uploaded_price, platform = gmarket(soup)     
                        
                        
                elif "옥션" in html :    
                    seller, uploaded_price, platform = auction(soup)     
                        
                        
                elif "ssg.com" in html :    
                    seller, uploaded_price, platform = ssg(soup)     
                        
                elif "11번가" in html :        
                    seller, uploaded_price, platform = elevenst(soup)     
                        
                        
                elif "위메프" in html :        
                    seller, uploaded_price, platform = wemakeprice(soup)     
                        
                elif "lotte" in html :       
                    seller, uploaded_price, platform = lotte(soup)                          
                        
                else :       
                    seller, uploaded_price, platform = unknown(soup)  
        
                ddf_total.loc[i,'판매처'] = seller
                ddf_total.loc[i,'플랫폼'] = platform   
                
                
    except AttributeError:       
         pass       
       
    except TypeError:       
        pass       
           
    except TimeoutException :       
        pass       
    except UnexpectedAlertPresentException: 
        pass 
            
    # '가격지도' 폴더 경로 생성
    dir_path = os.path.join('C:\\', '가격지도')
    # '가격지도' 폴더가 없으면 생성
    if not os.path.exists(dir_path):
        os.mkdir(dir_path)
    # 현재 시간으로 파일명 생성
    now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f'가격지도_{now}.xlsx'
    # 파일 경로 생성
    file_path = os.path.join(dir_path, filename)
    # 파일 저장
    ddf_total.to_excel(file_path, index=False)
    print("작업이 완료되었습니다.")
            
      
def task3():     
         
    링크 = []     
    표시단가 = []     
         
    # 엑셀 파일 읽어오기       
    file_path = filedialog.askopenfilename(title="가격지도 파일", defaultextension=".xlsx")       
    df_input = pd.read_excel(file_path, header =0)      
         
    df_input.head()     
         
    options = Options()       
    options = webdriver.ChromeOptions()       
    driver = webdriver.Chrome('./mnt/c/chromedriver.exe', options=options)       
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": """ Object.defineProperty(navigator, 'webdriver', { get: () => undefined }) """})       
         
         
         
    for i in tqdm(range(len(df_input))):    
        try :    
            url = str(df_input.loc[i, '링크주소'])       
            driver.get(url)     
            time.sleep(8)     
            html = driver.page_source       
            soup = BeautifulSoup(html, 'lxml')       
            time.sleep(1)     
                 
            if "NAVER" in html :       
                seller, uploaded_price, platform = naver(soup)     
                 
                 
            elif "coupang" in html :       
                seller, uploaded_price, platform = coupang(soup)     
                 
                 
            elif '인터파크쇼핑' in html :                                      
                seller, uploaded_price, platform = interpark(soup)     
                 
                 
            elif "G마켓" in html :       
                seller, uploaded_price, platform = gmarket(soup)     
                 
                 
            elif "옥션" in html :       
                seller, uploaded_price, platform = auction(soup)     
                 
                 
            elif "ssg.com" in html :       
                seller, uploaded_price, platform = ssg(soup)     
                 
            elif "11번가" in html :        
                seller, uploaded_price, platform = elevenst(soup)     
                 
                 
            elif "위메프" in html :        
                seller, uploaded_price, platform = wemakeprice(soup)     
                 
            elif "lotte" in html :       
                seller, uploaded_price, platform = lotte(soup)                          
                 
            else :       
                seller, uploaded_price, platform = unknown(soup)     
                     
            링크.append(url)     
            표시단가.append(uploaded_price)     
        except UnexpectedAlertPresentException: 
            pass      
             
    df = pd.DataFrame({       
        "링크": 링크,     
        "표시단가" : 표시단가})     
         
         
    df_input['등록단가'] = df_input['링크주소'].map(df.set_index('링크')['표시단가'])     
    # '가격지도' 폴더 경로 생성        
    dir_path = os.path.join('C:\\', '가격지도')        
            
    # '가격지도' 폴더가 없으면 생성        
    if not os.path.exists(dir_path):        
        os.mkdir(dir_path)        
            
    # 현재 시간으로 파일명 생성        
    now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")        
    filename = f'가격지도_단가확인{now}.xlsx'        
            
    # 파일 경로 생성        
    file_path2 = os.path.join(dir_path, filename)        
            
    # 파일 저장        
    df_input.to_excel(file_path2, index=False)      
         
    driver.close     
                
    print("3번 작업이 완료되었습니다.")       
        
        
             
# 도스 명령창 출력      
def main() :      
    while True:      
        os.system('cls' if os.name == 'nt' else 'clear')     
             
        print("가격지도 관리 프로그램")      
        print("1. 개별작업진행")      
        print("2. 엑셀 읽어오기")      
        print("3. 엑셀 가격정보 가져오기")     
        print("    ⊂_ヽ")
        print("     ＼＼ Λ＿Λ")
        print("      ＼( ‘ㅅ’ ) 두둠칫")
        print("         > ⌒ヽ")
        print("       /   へ＼")
        print("      /    / ＼＼")
        print("    ﾚ ノ    ヽ_つ")
        print("   / /두둠칫")
        print("   / /|")
        print("   ( (ヽ")
        print(" 　|　|、＼")
        print(" 　| 丿 ＼ ⌒)")
        print("  | |　　) /")
        print("  `ノ )    L/")
       

        # 사용자 선택 입력 받기      
        user_input = input()      
        # 선택에 따라 작업 실행      
        if user_input == "1":      
            task1()     
        elif user_input == "2":      
            task2()     
        elif user_input == "3":      
            task3()     
        else:      
            print("잘못된 입력입니다.")   
               
               
                 
if __name__ == "__main__":     
    main()
