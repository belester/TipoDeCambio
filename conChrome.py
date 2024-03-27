from selenium import webdriver

def conChromeDriver():

    def __init__(self, download_dirpath=None, visibility=False):
        opts = webdriver.ChromeOptions()

        if download_dirpath:
            prefs = {'download.default_directory': download_dirpath}
            opts.add_experimental_option('prefs', prefs)

        opts.add_argument('--disable-extensions')

        if not visibility:
            opts.add_experimental_option('excludeSwitches',
                                         ['ignore-certificate-errors'])
            opts.add_argument('--disable-gpu')
            # opts.add_argument('--headless')

        # Aquí directamente especificas la ruta al chromedriver
        path_to_chromedriver = '/chrome-win64/chrome.exe'

        super().__init__(
            webdriver.Chrome(executable_path=path_to_chromedriver, options=opts)
        )
