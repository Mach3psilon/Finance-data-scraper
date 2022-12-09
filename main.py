from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import warnings
from functools import wraps
import time
from pathlib import Path
import os
import xlsxwriter



#decorator function to collect time data 
def timeit(func):
    @wraps(func)
    def timeit_wrapper(*args, **kwargs):
        start_time = time.perf_counter()
        result = func(*args, **kwargs)
        end_time = time.perf_counter()
        total_time = end_time - start_time
        print(f'Successfully finished in {total_time:.4f} seconds...')
        return result
    return timeit_wrapper

#for xlsx sheet name
def get_time():
    now = time.localtime()
    return time.strftime("%m-%d-%Y", now)

#to get first day of the month
def get_first_day_of_month():
    now = time.localtime()
    return time.strftime("01/%m/%Y", now)

#to get today's date
def get_today():
    now = time.localtime()
    return time.strftime("%d/%m/%Y", now)

#to activate and adjust driver beforehand
def activate_driver():
    global driver
    chrome_options = Options()

    #remove it if you want to see warnings
    warnings.filterwarnings("ignore", category=DeprecationWarning) 
    chrome_options.add_argument('log-level=3')
    
    #remove it if you want to see what is going on at the background
    chrome_options.add_argument("--headless")
    
    #to download files to the specified directory
    prefs = {'download.default_directory' : str(Path.cwd())+'\Currencies'}
    chrome_options.add_experimental_option('prefs', prefs)
    
    driver = webdriver.Chrome('./chromedriver.exe', chrome_options=chrome_options)
    
def get_currency_data():
    
    driver.get("https://finance.yahoo.com/currencies")
    result_list = []
    
    i = 0
    time.sleep(2)
    try:
        while True:
            i += 1
            
            append_list = []
            name = driver.find_element(By.XPATH, '//*[@id="list-res-table"]/div[1]/table/tbody/tr[' + str(i) + ']/td[1]/a').text
            append_list.append(name[:-2:])
            last_price = driver.find_element(By.XPATH, '//*[@id="list-res-table"]/div[1]/table/tbody/tr[' + str(i) + ']/td[2]').text
            append_list.append(last_price)
            for j in range(3):
                row_element = driver.find_element(By.XPATH, '//*[@id="list-res-table"]/div[1]/table/tbody/tr[' + str(i) + ']/td[' + str(j + 3) + ']/fin-streamer').text
                append_list.append(row_element)
            
            result_list.append(append_list)
    except:
        return result_list




def get_path():
    path = str(Path.cwd())
    print('--------------------------------------------')
    if os.path.exists(path + '/Currencies'):
        print("Currencies path exists")
    else:
        os.mkdir(path +'/Currencies')
        print('Currencies path created')
    return path 

    
def write_xlsx(data):
    path = get_path()
    workbook = xlsxwriter.Workbook(path + '/Currencies/Currencies.xlsx')
    worksheet = workbook.add_worksheet(str(get_time()))
    

    #three color shift in xlsx writer wasn't working properly so I had to do it manually
    data_val = data[0][4]
    
    max = float(data_val.replace('%', ''))
    min = float(data_val.replace('%', ''))
    for i in range(len(data)):
        data_val = data[i][4]
        if float(data_val.replace('%', '')) > max:
            max = float(data_val.replace('%', ''))
        elif float(data_val.replace('%', '')) < min:
            min = float(data_val.replace('%', ''))
    
    gap = max - min
   
    def format(x, gap):
        
        return float(str((1 / gap) * x)[0:5])
    
    cell_format_yellow = workbook.add_format({'align': 'right'})
    cell_format_yellow.set_bg_color('#F1EB9C')
    cell_format_green = workbook.add_format({'align': 'right'})
    cell_format_green.set_bg_color('#AFC39E')
    cell_format_red = workbook.add_format({'align': 'right'})
    cell_format_red.set_bg_color('#D99888')
    cell_format_red_bright = workbook.add_format({'align': 'right'})
    cell_format_red_bright.set_bg_color('#FF5733')
    cell_format_green_bright = workbook.add_format({'align': 'right'})
    cell_format_green_bright.set_bg_color('#006400')
    #----------------------------------------------------------------------------------
    
    
    

    bold = workbook.add_format({'bold': True})
    align_right = workbook.add_format({'align': 'right'})
    align_left = workbook.add_format({'align': 'left'})
    worksheet.set_column(0, 4, 10)
    worksheet.write('A1', 'Symbol', bold)
    worksheet.write('B1', 'Name', bold)
    worksheet.write('C1', 'Last Price', bold)
    worksheet.write('D1', 'Change', bold)
    worksheet.write('E1', '% Change', bold)
    for i in range(len(data)):
        for j in range(len(data[i])):
            if j == 4:
                formatted_data = data[i][j]
                formatted_data = formatted_data.replace('%', '')
                formatted_data = float(formatted_data.replace('+', ''))
                formatted_data = format(formatted_data, gap)
                
                
                if formatted_data == 0:
                    worksheet.write(i + 1, j, data[i][j], cell_format_yellow)
                elif formatted_data > 0.5:
                    worksheet.write(i + 1, j, data[i][j], cell_format_green_bright)
                elif 0 < formatted_data < 0.5:
                    worksheet.write(i + 1, j, data[i][j], cell_format_green)
                elif -0.5 < formatted_data < 0:
                    worksheet.write(i + 1, j, data[i][j], cell_format_red)
                elif formatted_data < -0.5:
                    worksheet.write(i + 1, j, data[i][j], cell_format_red_bright)

                
                

                
                
                
                
                
            elif 4 > j > 1:
                worksheet.write(i + 1, j, data[i][j], align_right)
            else:

                worksheet.write(i + 1, j, data[i][j], align_left)
    print('Currencies.xlsx created\n')
    workbook.close()



def download_top_five_currencies():
    print('--------------------------------------------')
    print('Downloading currencies...')
    def download_currency(index):
        driver.get("https://finance.yahoo.com/currencies")
        driver.find_element(By.XPATH, '//*[@id="list-res-table"]/div[1]/table/thead/tr/th[5]').click()
        time.sleep(2)
        driver.find_element(By.XPATH, '//*[@id="list-res-table"]/div[1]/table/tbody/tr[' + index + ']/td[1]/a').click()
        time.sleep(2)

        #to supress pop-up
        try:
            driver.find_element(By.XPATH, '//*[@id="myLightboxContainer"]/section/button[1]').click()
        except: 
            pass
        time.sleep(2)
        driver.find_element(By.XPATH, '//*[@id="quote-nav"]/ul/li[4]/a').click()
        time.sleep(2)
        driver.find_element(By.XPATH, '//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div/div/div').click()
        start_date = driver.find_element(By.XPATH, '//*[@id="dropdown-menu"]/div/div[1]/input')
        start_date.send_keys(get_first_day_of_month())
        end_date = driver.find_element(By.XPATH, '//*[@id="dropdown-menu"]/div/div[2]/input')
        end_date.send_keys(get_today())
        driver.find_element(By.XPATH, '//*[@id="dropdown-menu"]/div/div[3]/button[1]').click()
        
        driver.find_element(By.XPATH, '//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/button').click()
        driver.find_element(By.XPATH, '//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[2]/span[2]/a').click()
        time.sleep(2)

    #change it to see more currencies.
    currency_range = 5
    for i in range(currency_range):
        #Changing to url twice because of some weird bug that the site has.
        driver.get("https://finance.yahoo.com/currencies")
        download_currency(str(i + 1))
        print('Currency ' + str(i + 1) + '/' + str(currency_range) + ' downloaded')
    
    
    
@timeit
def main():
    print('--------------------------------------------')
    print('Starting program')
    activate_driver()
    write_xlsx(get_currency_data())
    download_top_five_currencies()
    driver.close()

if __name__ == "__main__":
    main()
    
    
    