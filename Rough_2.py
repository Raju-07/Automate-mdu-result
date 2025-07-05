# Required framework and library
import time
from tkinter import messagebox as msg
#"Trying to automate the task using selenium"
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options


# Set up the WebDriver (Edge in this example)
options = Options()
options.add_experimental_option('excludeSwitches', ['enable-logging'])  # Suppress DevTools message
driver = webdriver.Edge(options=options)

# Navigate to the result portal
driver.get("https://result.mdu.ac.in/postexam/result.aspx")

res_std_data = []

# Pause to ensure the page loads fully
time.sleep(1)

#"Enter Your data Here 👇👇👇"
#Data format (registration,Roll No)
student_details = [(2312051470,6071607)]

# This data 👆👆 is just an example enter your data properly or take Refferance view of the requirement.txt file which is also in this Repo.

for Registration_No,Roll_no in student_details:
    try:

        # Locate the Registration No and Roll No fields and enter the data
        driver.find_element(By.ID, "txtRegistrationNo").send_keys(Registration_No)  # Registration Number one by one
        driver.find_element(By.ID, "txtRollNo").send_keys(Roll_no)  # Roll Number one by one

        # Submit the form
        driver.find_element(By.ID, "cmdbtnProceed").click()

        # Pause to let the results load
        time.sleep(3)
        driver.find_element(By.ID,"imgComfirm").click()
        time.sleep(4)
 # Locate the element view and click on it
        view_link = driver.find_element(By.ID, "rptMain_ctl01_lnkView") 
        view_link.click()

        time.sleep(5)

#Extracting all the required Details from the result
        student_name = driver.find_element(By.ID,"lblStudentName").text
        time.sleep(1)
        father_name = driver.find_element(By.ID,"lblFatherName").text
        #Extracting Subjects Marks 
        bca201 = driver.find_element(By.XPATH,"//table//tr[2]//td[8]").text
        bca202 = driver.find_element(By.XPATH,"//table//tr[3]//td[8]").text
        bca203 = driver.find_element(By.XPATH,"//table//tr[4]//td[8]").text
        bca204 = driver.find_element(By.XPATH,"//table//tr[5]//td[8]").text
        bca205 = driver.find_element(By.XPATH,"//table//tr[6]//td[8]").text
        #Result
        total_marks = driver.find_element(By.ID,"rptMarks_ctl06_lblTotal").text
        result = driver.find_element(By.ID,"lblresult").text

        print(f"{student_name = }")
        print(f"{father_name = }")
        print(f"{Registration_No = }")
        print(f"{Roll_no = }")
        
        print(f"{bca201 = }")
        print(f"{bca202 = }")
        print(f"{bca203 = }")
        print(f"{bca204 = }")
        print(f"{bca205 = }")
        print(f"{total_marks = }")
        print(f"{result = }")

    except:
        pass