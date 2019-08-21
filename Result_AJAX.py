#Import all necessary modules
import re 
import os
import sys
import time
import xlsxwriter
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


#Start the driver
'''driver = webdriver.Chrome('C:/Users/Pranav/Downloads/chromedriver.exe')
'''
#Open chrome in headless mode i.e. without actually opening the window
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument("--silent")
driver = webdriver.Chrome('C:/Users/Pranav/Downloads/chromedriver.exe', chrome_options=chrome_options)

#Initialization of variable and a list
row=0
usn=[]

#Input the Batch 
batch = input('Enter the batch (15/16/17/18): ')

#Input the Department
dept = input('Enter the department (cs/is/ec/me/te/cv): ')

#Input the Starting USN (Number format, Ex:1)
usn_start = int(input('Enter Starting USN: '))

#Input the Ending USN (Number format, Ex:115)
usn_end = int(input('Enter Starting USN: '))

#Input Subject Code (String format, Ex:15cs44)
subject_code = input('Enter Subject Code : ')
subject_code = subject_code.upper()

#Extracting the semester from the subject code
sem = subject_code[-2:-1]

#Condition for non-diploma USNs
if usn_start < usn_end and usn_end < 400:
       
    u = '1by'+batch+dept

    #Creating a folder
    folder_name=u
    u_upper=u.upper()
    a=os.getcwd()
    if not os.path.exists(folder_name):
        os.mkdir(folder_name)
        a=a+'\\'+folder_name
        os.chdir(a)
        print('Folder Created')

    #Create and initialize the workbook
    workbook = xlsxwriter.Workbook(u_upper+'_'+subject_code+'.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write(row,0,"Name")
    worksheet.write(row,1,"USN")
    worksheet.write(row,2,"Subject")
    worksheet.write(row,3,"Grade")
    worksheet.write(row,4,"Credits")
    worksheet.write(row,5,"Internal Marks")
    worksheet.write(row,6,"External Marks")
    worksheet.write(row,7,"Total Marks")
    worksheet.set_column(0, 7, 40)

    #Generate USNs for the given range
    for i in range(usn_start,usn_end+1):
        usn.append(u+format(i,'03d'))

    #Start the fetching process
    row = 1
    print('Fetching')

    #Loop for the given USN Range 
    for i in range((usn_end-usn_start)+1):
        
        #Try-catch for better exception handling
        try:
            
            #Generate the URL and open that using the webdriver
            url = "https://www.vtu4u.com/result/"+usn[i]+"/sem-"+sem+"/rs-22?cbse=1"
            driver.get(url)                
            time.sleep(1)

            #Parse the source using bs4       
            soup = BeautifulSoup(driver.page_source,'lxml')
            
            sc = soup.find(text=subject_code)
            
            #Subject Code
            name = sc.findParent().findNextSibling()
              
            #Internal Marks               
            internal = name.findNextSibling()
            
            #External Marks       
            external = internal.findNextSibling()
            
            #Total Marks       
            total = external.findNextSibling()

            #Credits
            credit = total.findNextSibling().findNextSibling()
            
            #Grade        
            grade = total.findNextSibling().findNextSibling().findNextSibling()
            
            #Name
            std = soup.find('div',{'class':'student_details'}).p.contents[2]
            name1 = str(std)
            name1 = re.sub('[\xa0]', '', name1)
            fname = name1.strip()

            #Write the data into the sheet
            worksheet.write(row,0,fname)
            worksheet.write(row,1,usn[i].upper())
            worksheet.write(row,2,subject_code)
            worksheet.write(row,3,grade.getText())
            worksheet.write(row,4,credit.getText())
            worksheet.write(row,5,internal.getText())
            worksheet.write(row,6,external.getText())
            worksheet.write(row,7,total.getText())
            
            time.sleep(1)       
            
        except:
            pass
        
        finally:
            row = row+1
            
    #Close and save the workbook
    workbook.close()
    print('Done')

#Condition for diploma USNs
elif usn_start < usn_end and usn_end >= 400:

    #Slight modifications for catering to diploma USNs
    batch1=int(batch)
    batch1=batch1+1
    batch2=str(batch1)

    #Folder creation
    u = '1by'+batch2+dept
    folder_name=u
    u_upper=u.upper()
    
    a=os.getcwd()
    if not os.path.exists(folder_name):
        os.mkdir(folder_name)
        a=a+'\\'+folder_name
        os.chdir(a)
        print('Folder Created')

    #Create and initialize the workbook
    workbook = xlsxwriter.Workbook('DIP_'+batch+'_'+u_upper+'_'+subject_code+'.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write(row,0,"Name")
    worksheet.write(row,1,"USN")
    worksheet.write(row,2,"Subject")
    worksheet.write(row,3,"Grade")
    worksheet.write(row,4,"Credits")
    worksheet.write(row,5,"Internal Marks")
    worksheet.write(row,6,"External Marks")
    worksheet.write(row,7,"Total Marks")
    worksheet.set_column(0, 7, 40)

    #Generate USNs
    for i in range(usn_start,usn_end+1):
        usn.append(u+format(i,'03d'))

    #Start the fetching process
    row = 1
    print('Fetching')

    #Loop for the given USN Range
    for i in range((usn_end-usn_start)+1):

        #Try-catch for better exception handling
        try:
            
            #Generate the URL and open that using the webdriver
            url = "https://www.vtu4u.com/result/"+usn[i]+"/sem-"+sem+"/rs-19?cbse=1"
            driver.get(url)                
            time.sleep(1)

            #Parse the source using bs4       
            soup = BeautifulSoup(driver.page_source,'lxml')
            
            sc = soup.find(text=subject_code)

            #Subject Code
            name = sc.findParent().findNextSibling()
              
            #Internal Marks               
            internal = name.findNextSibling()
            
            #External Marks       
            external = internal.findNextSibling()
            
            #Total Marks       
            total = external.findNextSibling()

            #Credits
            credit = total.findNextSibling().findNextSibling()
            
            #Grade        
            grade = total.findNextSibling().findNextSibling().findNextSibling()
            
            #Name
            std = soup.find('div',{'class':'student_details'}).p.contents[2]
            name1 = str(std)
            name1 = re.sub('[\xa0]', '', name1)
            fname = name1.strip()

            #Write the data into the sheet
            worksheet.write(row,0,fname)
            worksheet.write(row,1,usn[i].upper())
            worksheet.write(row,2,subject_code)
            worksheet.write(row,3,grade.getText())
            worksheet.write(row,3,credit.getText())
            worksheet.write(row,5,internal.getText())
            worksheet.write(row,6,external.getText())
            worksheet.write(row,7,total.getText())
            
            time.sleep(1)       
            
        except:
            pass
        
        finally:
            row = row+1
            
    #Close and save the workbook
    workbook.close()
    print('Done')

#Release the driver
driver.close()
driver.quit()
sys.exit()


