from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import chromedriver_autoinstaller
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import logging

#configuring logging object
logging.basicConfig(filename="newfile.log",
                    format='%(asctime)s %(message)s',
                    filemode='w')
#Creating an object
logger=logging.getLogger()
  
#Setting the threshold of logger to DEBUG
logger.setLevel(logging.DEBUG)



#method for creating chrome driver
def create_driver_object(request_url):
#     setting the option for driver
    options = Options()
    options.add_argument('--headless')

    #installing chrome driver for first time
    chromedriver_autoinstaller.install()
    driver = webdriver.Chrome(service=Service())
    driver.get(request_url)
    driver.implicitly_wait(10)
    return driver
    
    
#method for creating excel file object
def create_excel_file(name=None, column_name_lst=[]):
    if name and column_name_lst:
        import xlsxwriter
        workbook = xlsxwriter.Workbook('{}.xlsx'.format(name))
        worksheet = workbook.add_worksheet("My sheet")
        col = 0
        for each in column_name_lst:
            worksheet.write(0, col, each)
            col = col+1
    
        return worksheet, workbook
    else:
        return None, None
        
        
def extracting_data(search_data=None):
    worksheet, workbook = create_excel_file(name ="yello_page_{}".format(search_data.replace(' ', '_')), column_name_lst=['Name','Contact Number','Address','Type','Official Mail'] )
    page_no = 1
    while True:
        driver = create_driver_object('https://www.yellowpages.com.au/search/listings?clue={}&pageNumber={}'.format(search_data,page_no))
        all_data = driver.find_elements_by_xpath("//div[@class='Box__Div-dws99b-0 bMUVdR MuiPaper-root MuiCard-root MuiPaper-elevation1 MuiPaper-rounded']")
        # print(type(all_data))    
        #iterating the object
        if workbook:
            for i in range(len(all_data)):
                try:
                    col = 0
                    if driver.find_elements_by_xpath("//div[@class='Box__Div-dws99b-0 jbEoDe' and text()='More info']"):
                        driver.find_elements_by_xpath("//div[@class='Box__Div-dws99b-0 jbEoDe' and text()='More info']")[i].click()
            
                    name = driver.find_elements_by_xpath("//a[@class='MuiTypography-root MuiLink-root MuiLink-underlineNone MuiTypography-colorPrimary']")[i].text
                    number = driver.find_elements_by_xpath("//button[@class='MuiButtonBase-root MuiButton-root MuiButton-text MuiButton-textPrimary MuiButton-fullWidth']")[i].text
                    address = driver.find_elements_by_xpath("//div[@class='Box__Div-dws99b-0 dzJNWw']//p")[i].text
                    next_url = driver.find_elements_by_xpath("//div[@class='Box__Div-dws99b-0 bMUVdR MuiPaper-root MuiCard-root MuiPaper-elevation1 MuiPaper-rounded']//a[@class='MuiTypography-root MuiLink-root MuiLink-underlineNone MuiTypography-colorPrimary']")[i].get_attribute('href')
                    new_driver = create_driver_object(next_url)
                    new_txt = new_driver.find_element_by_xpath("//a[@class='contact contact-main contact-email']").get_attribute("data-email")
                    print(new_txt)
                    if worksheet:
                        worksheet.write(i+1, col, name)
                        worksheet.write(i+1, col + 1, number)
                        worksheet.write(i+1, col + 2, address)
                        worksheet.write(i+1, col + 3, search_data)
                        if new_txt:
                            worksheet.write(i+1, col + 4, new_txt)
                        else:
                            worksheet.write(i+1, col + 4, 'Not availabel')
                            
        
                    print(name,'-------', number, '--------',address, '-----------',next_url)
                    new_driver.close()
                except Exception as e:
                    logger.error(e)
                    workbook.close()
                    pass
        page_no = page_no + 1
        break


    workbook.close()
    
    
    
    
#function calling
extracting_data('Creche')
