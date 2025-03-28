# Required framework and library
import pyautogui
from tkinter import messagebox as msg
"Trying to automate the task using selenium"
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Set up the WebDriver (Edge in this example)
options = Options()
options.add_experimental_option('excludeSwitches', ['enable-logging'])  # Suppress DevTools message
driver = webdriver.Chrome(options=options)


# Navigate to the result portal
driver.get("https://result.mdu.ac.in/postexam/result.aspx")

student_details = []   #List for storing the student details

# Pause to ensure the page loads fully
pyautogui.sleep(1)

"Enter Your data Here 👇👇👇"
students_detail = [
    ("Registration No","Roll NO"),
    (2312051573, 6071620)
     ] 

# This data 👆👆 is just an example enter your data properly or take Refferance view of the requirement.txt file which is also in this Repo.

for Registration_No,Roll_no in students_detail:
    try:

        # Locate the Registration No and Roll No fields and enter the data
        driver.find_element(By.ID, "txtRegistrationNo").send_keys(Registration_No)  # Replace with your Registration No
        driver.find_element(By.ID, "txtRollNo").send_keys(Roll_no)  # Replace with your Roll No

        # Submit the form
        driver.find_element(By.ID, "cmdbtnProceed").click()

        # Pause to let the results load
        pyautogui.sleep(3)
        driver.find_element(By.ID,"imgComfirm").click()
        pyautogui.sleep(4)

        view_link = driver.find_element(By.ID, "rptMain_ctl01_lnkView")  # Locate the element view and click on it
        view_link.click()

        pyautogui.sleep(5)

      
        student_name = driver.find_element(By.ID,"lblStudentName").text
        pyautogui.sleep(1)

        "Getting all data from the result window"
        father_name = driver.find_element(By.ID,"lblFatherName").text
        # os_marks = driver.find_element(By.XPATH,"//table//tr[2]//td[8]").text   "Used for getting the data from the result Table"
  
        "Finding Total Marks"
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

        student_details.append(data)
        
        pyautogui.sleep(1)
        driver.refresh()
        pyautogui.sleep(2)

    except Exception as e:
        print(f"Error is occured In {Registration_No} and \nError is {e}")
        msg.showerror("process Interupted","Unexpected Error: Please ensure good internet connect or try again")

print(student_details)
print()   # Giving Extra Space for better Representation
print()

"Representing Data in well format"
print(f"{'Name':<22}: {'Total Marks':<13} {'Reappear ':^20}")
for dictionary in student_details:
    print(f"{dictionary['Name']:<22}: {dictionary['Total Marks']:<13}| {dictionary['Re appear ']:<20}")

# Display info when process complete
msg.showinfo("Success","All the data fatched successfully")

# Close the browser
driver.quit()
