from selenium import webdriver

def conCromeDriver():

    #options = webdriver.ChromeOptions()
    #driver = webdriver.Chrome(executable_path = 'C:\\Users\\A2278458\\Downloads\\TipoDeCambio\\chromedriver-win64\\chromedriver.exe', options=options)

    driver = webdriver.Chrome()
    return driver