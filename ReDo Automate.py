import webbrowser  # webbrowser is used for Opening the browser
import pyautogui as pag   # this lib is used for 
from tkinter import messagebox as msg


class Automate:
    # Initializing the Constructor to open the result portal by default when program execute.
    def __init__(self):
        URL = "https://result.mdu.ac.in/postexam/result.aspx" #URL of the result portal of the MDU.
        webbrowser.open_new(URL)


    def fill_info(self):
        
        # Waiting for a little movement to load the site successfully
        pag.sleep(5)
        #List for storing the students details to downlaod the result
        students_detail = [(2312051470,6071607,8076954807,"rajuyadav782760@gmail.com"),(2312051471,6071611,9115004409,"ritikyadav4409@gmail.com"),(2312051513,6071579,7982538645,"yadavnitin24821@gmail.com")]

        
        # Starting a for loop which will fill all the details and 

        for Registraion_no,Roll_no,Phone_no,Email_id in students_detail:

            #clicking on the registration Entrybox widget
            pag.click(x=549,y=143) #Coordination of the screen where the entrybox appear it depends on the resolution of the screen
            pag.write(str(Registraion_no))
            pag.sleep(0.3)
            pag.press('tab')
            pag.write(str(Roll_no))
            pag.sleep(0.3)
            pag.press('enter')

            pag.sleep(2)

            #Creating a window for asking if we need to fill the other details too like phone no and email 
            more_detail=msg.askquestion(title="More Details",message="Do you need more details for result")
            if more_detail == "yes":
                # Write code here for more infomation
                raise NotImplementedError
            
            else:
                self.after_detail(Registraion_no)
    


    def after_detail(self,regis_no):
                self.registration_no = regis_no
        # Procceding further if you don't need additional data
                pag.moveTo(945,194,0.4)
                pag.sleep(0.2)
                pag.click()# pressing on the 'confirm' button
                pag.sleep(0.6)
                pag.moveTo(1018,170) # visiting to the coordinate where 
                pag.sleep(3) # wait for Response
                pag.leftClick(1018,172) # Pressing on the 'view' hyper link
                pag.sleep(5)
                # pag.doubleClick(537,435)
                # pag.hotkey('ctrl','c')  # for copying the student name
                pag.click(959,221) # Pressing on the 'print' button
                pag.sleep(2.5)
                pag.moveTo(189,817) # pessing save button
                pag.click()  # pressing on the 'save' button
                pag.sleep(1)
                # pag.hotkey('ctrl','v')

                pag.write(f"{self.registration_no} 3rd sem result")
                # pag.press('enter')
                pag.sleep(3)
                pag.press('enter')
                pag.sleep(2)


                pag.click(70,50)   # clicking on the refresh button
                pag.sleep(2)


if __name__ == "__main__":
    obj = Automate()
    obj.fill_info()
