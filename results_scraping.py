import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import os
from bs4 import BeautifulSoup as bs

def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)
folder = input("Enter the Folder Name to store the RESULT :  ")
createFolder(f'./{folder}/')


def pd_to_xlsx():
            df = pd.DataFrame.from_dict(columns, orient='index')
            df = df.transpose()
            df.to_excel("BE CSE II YEAR RESULT.xlsx")
            print("The file written successfully . ")



def result_page():
    driver = webdriver.Chrome()
    
    file_name = input("Save the file NAME as(applies to all files) : ")
    for i in  range(201008001,201008005):
        try:
            driver = webdriver.Chrome("C:\\Users\\Lenovo\\Documents\\py\\selenim\\chromedriver.exe")
            driver.get("https://coe.annamalaiuniversity.ac.in/rgl_result.php")
            driver.find_element("name","item").send_keys("ENGG")
            driver.find_element("name","txtrollregno").send_keys(i)
            driver.find_element("name","Submit").click()
            num = driver.find_element(By.CLASS_NAME,"table")
            s = num.find_element(By.CLASS_NAME,"table-bordered").text
            soup = bs(s,'html.parser')
            with open(f"{os.getcwd()}\\{folder}\\{file_name}{i}.txt",'w') as f:
                f.write(str(soup))
            read = open(f"{os.getcwd()}\\{folder}\\{file_name}{i}","r")
            raw_data = []
            for i in read:
                raw_data.append(i.split())
                de = raw_data[:4]
                sc =raw_data[4:]
            columns = {" ".join(de[0][0:2]) : []," ".join(de[1][0:1]) : []," ".join(de[2][0:1]) : [],"".join(de[3][0]) : [],"-".join(sc[0][1:4]) : [],
                           " ".join(sc[1][1:6]) : []," ".join(sc[2][1:5]) : []," ".join(sc[3][1:5]) : []," ".join(sc[4][1:7]) : [], " ".join(sc[5][1:5]) : [], " ".join(sc[6][1:4]) : [],
                           " ".join(sc[7][1:4]) : []," ".join(sc[8][1:6]) : [],"".join(de[3][-1]) : []}
            columns[" ".join(de[0][0:2])].append("".join(de[0][3]))
            columns[" ".join(de[1][0:1])].append(" ".join(de[1][2:]))
            columns[" ".join(de[2][0:1])].append(" ".join(de[2][2:]))
            columns["".join(de[3][0]) ].append(" ".join(sc[0][0]))
            columns["-".join(sc[0][1:4])].append("".join(sc[0][6]))
            columns[" ".join(sc[1][1:6])].append(" ".join(sc[1][8]))
            columns[" ".join(sc[2][1:5])].append("".join(sc[2][7]))
            columns[" ".join(sc[3][1:5])].append("".join(sc[3][7]))
            columns[" ".join(sc[4][1:7])].append("".join(sc[4][9]))
            columns[" ".join(sc[5][1:5])].append("".join(sc[5][7]))
            columns[" ".join(sc[6][1:4])].append("".join(sc[6][6]))
            columns[" ".join(sc[7][1:4]) ].append("".join(sc[7][6]))
            columns[" ".join(sc[8][1:6])].append("".join(sc[8][8]))
            columns["".join(de[3][-1])].append("".join(sc[8][-1]))
            pd_to_xlsx()

   
            
        except:
            pass
result_page()
        



        
        



