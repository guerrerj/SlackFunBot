import os 
from subprocess import Popen, PIPE 
from selenium import webdriver 
from selenium.webdriver import ActionChains
import time
import logging

""" This module contains logic for taking screenshot 
    Requirements: 
    * Selenium needs to be installed - pip install selenium 
    * Chromedriver needs to be downloaded and saved to path 
    * Need to download and install ImageMagick

"""
absPath = lambda *p: os.path.abspath(os.path.join(*p)) 
ROOT = absPath(os.path.dirname(__file__)) 

def executeCommand(command):
    """ Executes the screenshot command """
    result = Popen(command, shell=True, stdout=PIPE).stdout.read() 
    if len(result) > 0 and not result.isspace():
        raise Exception(f"Failed to execute command with result : {result}")
    
def performScreenCapture(url, screenPath, width, height, closeButton=False):
    """ Uses a webdriver to capture a screen shot of the webpage """
    driver = webdriver.Chrome() 
    # Log file will be in the same directory 
    driver.set_script_timeout(30) 
    if width and height: 
        driver.set_window_size(width, height) 
    driver.get(url) 
    time.sleep(5)
    if closeButton:
        try:
            button = driver.find_element_by_css_selector("a.white.pull-right.bolder")
            actions = ActionChains(driver)
            actions.move_to_element(button)
            actions.click(button).perform()
        except Exception as ex:
            logging.warning("Unable to close button")
        
    driver.execute_script("window.scrollTo(0,100)")
    
    driver.save_screenshot(screenPath)
    
def performCrop(params):
    """ Crops a screenshot to the indicated size in the params dictionary """
    command = ['convert', params['screenPath'], '-crop', 
               f'{params["width"]}x{params["height"]}+0+0', params['cropPath'] ]
    executeCommand(' '.join(command)) 
    
def getThumbNail(params):
    """ Creates a thumbnail image from the screen shot image """ 
    command = ['convert', params['cropPath'],  '-thumbnail', 
               f'{params["width"]}x{params["height"]}', params['thumbnailPath']]
    executeCommand(' '.join(command)) 
    
def getScreenShot(**kwargs):
    """ This is the main method that captures a screenshot """
    url = kwargs['url']
    width = int(kwargs.get('width', 1024))
    height = int(kwargs.get('height', 768))
    fileName = kwargs.get('fileName', 'screen.png') 
    path = kwargs.get('path', ROOT)
    closeButton = kwargs.get('closeButton')
    
    crop = kwargs.get('crop', False) 
    cropWidth = int(kwargs.get('cropWidth', width)) 
    cropHeight = int(kwargs.get('cropHeight', height))
    cropReplace = kwargs.get('cropReplace', False)
    
    thumbnail = kwargs.get('thumbnail', False) 
    thumbnailWidth = int(kwargs.get('thumbnailWidth', width)) 
    thumbnailHeight = int(kwargs.get('thumbnailHeight', height)) 
    thumbnailReplace = kwargs.get('thumbnailReplace', False) 
    
    screenPath = absPath(path, fileName) 
    cropPath = thumbnailPath = screenPath 
    
    if thumbnail and not crop: 
        raise Exception('Thumbnail generation requires a crop image')
    
    performScreenCapture(url, screenPath, width, height, closeButton)
    
    if crop:
        if not cropReplace:
            cropPath = absPath(path, 'crop_'+fileName)
        params = {'width': cropWidth, 'height': cropHeight, 
                    'cropPath':cropPath, 'screenPath': screenPath}
        performCrop(params)  
        if thumbnail:
            if not thumbnailReplace:
                thumbnailPath = absPath(path, 'thumbnail_'+fileName)
            params = {'width': thumbnailWidth, 'height': thumbnailHeight, 
                      'thumbnailPath': thumbnailPath, 'cropPath':cropPath}
            getThumbNail(params)
    return screenPath, cropPath, thumbnailPath

if __name__ == '__main__':
    url = 'https://www.merriam-webster.com/word-of-the-day'
    screenPath, cropPath, thumbnailPath = getScreenShot(url=url, fileName='test.png', crop=True, cropReplace=False, 
                                                        thumbnail=False, thumbnailReplace=False, thumbnailWidth=200, thumbnailHeight=150,closeButton=False)
    