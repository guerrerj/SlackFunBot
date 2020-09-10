import os 
import socket 
import win32serviceutil
import servicemanager
import win32event 
import win32service 
import slack
from slack import WebClient 
import time, sched 
import logging
import ScreenShot
from ScreenShot import getScreenShot
import sys
import traceback
#NB Service must be started as a specific user or else it won't start

# Set logging configuration 
logging.basicConfig(filename='debug.log', level=logging.DEBUG)
        
class WindowsService(win32serviceutil.ServiceFramework):
    """ This is a windows service class which will update slack with joke of the day
        and with a word of the day. 
    """
    _svc_name_ = "dailyKnowledge"
    _svc_display_name_ = "Daily Knowledge Bot"
    description = "Places a daily joke and word of the day from the internet to slack"
    isRunning = False
    scheduler = sched.scheduler(time.time, time.sleep)
    
    @classmethod 
    def parseCommandLine(cls):
        """ Parses the class in the command line """   
        win32serviceutil.HandleCommandLine(cls)
        
    def __init__(self, args):
        """ Class constructor - initializes the windows service """
        logging.debug("In Constructor\n")
        win32serviceutil.ServiceFramework.__init__(self,args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(200)  
        
    def SvcStop(self):
        """ Is a handler for when the service stops """
        logging.debug("Stopping service")
        print("Service is stopping")
        self.stop() 
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop) 
        
    def SvcDoRun(self):
        """ Is a handler for when a service is asked to start """
        logging.debug("Starting service")
        self.start() 
        #self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE, 
                              servicemanager.PYS_SERVICE_STARTED, 
                              (self._svc_name_, ''))
        try: # try main
            if not self.isRunning:
                logging.debug("Starting scheduler")
                #print("starting scheduler")   
                minutes = 1440           
                self.periodic(minutes*60, self.runAction) 
                #rint("starting scheduler")
                self.scheduler.run()
        except Exception as ex:
            print(ex)
            servicemanager.LogErrorMsg(traceback.format_exc()) # if error print it to event log
            os._exit(-1) #return some value other than 0 to os so that service knows to restart     
    
    
    def start(self):
        """ Set any additional starting conditions """
        self.isRunning = True  
    
    def stop(self):
        """ Set any additional stop conditions """ 
        self.isRunning = False 
    
    def periodic(self, interval, runAction):
        """ Schedules the screenshots to be taken at an interval """                
        if self.isRunning:
            logging.debug("Running main script")
            print("in periodic\n")
            
            print("entering schedule")
            self.scheduler.enter(interval, 1, self.periodic, 
                            (interval, runAction ))
            runAction()
        
    def runAction(self):
        """ Contains the main logic of the script """
        timeOut = 600 # 10 minutes till timeout 
        endTime = time.time() + timeOut   
        try:
            slackToken = os.environ["SLACK_API_TOKEN"]
            channelId = 'CGART6MC3'
            client = WebClient(token=slackToken)                             
            responseStatus = False
            screenPath1,_, _ = getScreenShot(url="https://www.ajokeaday.com/", 
                                            fileName='screenshot1.png', 
                                            crop=True, 
                                            cropReplace=False, 
                                            thumbnail=False, 
                                            thumbnailReplace=False, 
                                            thumbnailWidth=200, 
                                            thumbnailHeight=150, closeButton=True)
            
            screenPath2,_, _ = getScreenShot(url="https://www.merriam-webster.com/word-of-the-day", 
                                            fileName='screenshot2.png', 
                                            crop=True, 
                                            cropReplace=False, 
                                            thumbnail=False, 
                                            thumbnailReplace=False, 
                                            thumbnailWidth=200, 
                                            thumbnailHeight=150)
            
            while (not responseStatus and time.time() < endTime):
                screenPath1 = screenPath1.replace('\\', '/')
                screenPath2 = screenPath2.replace('\\', '/')                 
                
                
                response = client.chat_postMessage(
                    channel=channelId, icon_emoji=":exploding_head:", 
                        link_names=True, username=self._svc_display_name_,
                        blocks=[
                                {
                                    "type": "section",
                                    "text": {
                                        "type": "mrkdwn",
                                        "text": "@here Time to laugh and build your vocabulary!\n"
                                    }
                                }                                   
                            ])
                response2 = client.files_upload(channels=channelId, file=screenPath1, title="Joke of the day")
                response3 = client.files_upload(channels=channelId, file=screenPath2, title="Word of the day")
                if response["ok"] and response2["ok"] and response3["ok"]:
                    responseStatus = True
                else:
                    logging.warning(f"Error posting to channel: {response['error']}")
                time.sleep(5)
                
            if time.time() >= endTime:
                logging.warning("Timeout when posting to channel!")
                    
        except Exception as ex:
            logging.error(f"Error executing running service: {ex}")      
                         
                
                                        
    
if __name__ == '__main__':    
    WindowsService.parseCommandLine()