# Required framework and library
import time
from tkinter import messagebox as msg
#Automate the task using selenium"
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options


# Set up the WebDriver (Edge in this example)
options = Options()
options.add_experimental_option('excludeSwitches',['enable-logging'])  # Suppress DevTools message
driver = webdriver.Edge(options=options)
# Navigate to the result portal
driver.get("https://result.mdu.ac.in/postexam/result.aspx")
#Array to Store the data
res_std_data = []
# Pause to ensure the page loads fully
time.sleep(1)

#"Enter Your data Here 👇👇👇"
#Data format (registration,Roll No)
student_details = [(2312051612, 6071438), (2312051613, 6071439), (2312051507, 6071440), (2312051461, 6071441), (2312051801, 6071442), (2312051538, 6071568), (2312051455, 6071570), (2312051517, 6071571)]


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
        os_marks = driver.find_element(By.XPATH,"//table//tr[2]//td[8]").text  # "Used for getting the data from the result Table"
  
#        "Finding Total Marks"
        total_marks = driver.find_element(By.ID,"rptMarks_ctl06_lblTotal").text
        result = driver.find_element(By.ID,"lblresult").text
        reappear = ""
        if result == 'REP':
            re = driver.find_element(By.ID,"labelreapr").text
            reappear = reappear.join(re)

        data = {
                "Name":student_name,
                "Registraion No":Registration_No,
                "Roll No ": Roll_no,
                "Total Marks":total_marks,
                "Re appear ": reappear 
                    }

        res_std_data.append(data)
        
        time.sleep(1)
        driver.refresh()
        time.sleep(2)

    except Exception as e:
        print(f"Error is occured In {Registration_No} and \nError is {e}")
        msg.showerror("process Interupted","Unexpected Error: Please ensure good internet connect or try again")
        

print(res_std_data)
print()   # Giving Extra Space for better Representation

#"Representing Data in well format"
print(f"{'Name':<22}: {'Total Marks':<13} {'Reappear ':^20}")
for dictionary in res_std_data:
    print(f"{dictionary['Name']:<22}: {dictionary['Total Marks']:<13}| {dictionary['Re appear ']:<20}")

# Display info when process complete
msg.showinfo("Success","All the data fatched successfully")

# Close the browser
driver.quit()
