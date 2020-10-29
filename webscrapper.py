from selenium import webdriver
import xlwt 
import os
from xlwt import Workbook 
driver = webdriver.Chrome("C:\\Users\\Marcelo\\Desktop\\Data\\Chromedriver\\chromedriver")

url = "https://pt2.lamentosa.com/vampires-vs-werewolves-mmorpg-game/"


#inputs
email = input("Digite seu email\n")
senha = input("Digite sua senha\n")
nickname = input("Digite seu Nick\n")
print(" ")


#login
def site_login(url):
    global email
    global senha
    driver.get(url)
    driver.find_element_by_id("id_email").send_keys(email)
    driver.find_element_by_id ("id_password").send_keys(senha)
    driver.find_element_by_xpath("//div/div[2]/form/div/button").click()
    


    

#sheet header
wb = Workbook()

sheet = wb.add_sheet('Dados')
#Ally
sheet.write(0,8,"Inimigo")
sheet.write(0,0,"Aliado")
sheet.write(0,1,"Damage")
sheet.write(0,2,"Força")
sheet.write(0,3,"Defesa")
sheet.write(0,4,"Agilidade")
sheet.write(0,5,"Inteligencia")
sheet.write(0,6,"Resistencia")
sheet.write(0,7,"Level")
#Enemy
sheet.write(0,9,"Damage")
sheet.write(0,10,"Força")
sheet.write(0,11,"Defesa")
sheet.write(0,12,"Agilidade")
sheet.write(0,13,"Inteligencia")
sheet.write(0,14,"Resistencia")
sheet.write(0,15,"Level")
#Rows to write
row_A = 1
row_E = 1


#change page
def get_page(pn):
    try:
        data_url = "https://pt2.lamentosa.com/messages/list/attack/sender/"
        driver.get(data_url)
        first_link = "/html/body/section/div/div[3]/div[2]/table/tbody/tr[{}]/th/a"
        driver_url = driver.find_element_by_xpath(first_link.format(pn))
        driver_url = driver_url.get_attribute("href")
        driver.get(driver_url)
        driver_url = driver.find_element_by_xpath("/html/body/section/div/div[3]/div/div[2]/div/a[3]")
        driver_url = driver_url.get_attribute("href")
        driver.get(driver_url)
    except:
        None

      

        

#get combat info
def combat_info():
    global row_A
    global row_E
    global nickname
    
    #Attributes Ally
    
    forç = driver.find_element_by_xpath("/html/body/section[1]/div/div[2]/div[2]/ul[1]/li[2]/div/span[2]").text

    defe = driver.find_element_by_xpath("/html/body/section[1]/div/div[2]/div[2]/ul[1]/li[3]/div/span[2]").text

    agil = driver.find_element_by_xpath("/html/body/section[1]/div/div[2]/div[2]/ul[1]/li[4]/div/span[2]").text

    inte = driver.find_element_by_xpath("/html/body/section[1]/div/div[2]/div[2]/ul[1]/li[5]/div/span[2]").text

    resi = driver.find_element_by_xpath("/html/body/section[1]/div/div[2]/div[2]/ul[1]/li[6]/div/span[2]").text

    level = driver.find_element_by_xpath("/html/body/section[1]/div/div[2]/div[2]/div/div[1]/span[2]").text


    #Attribrutes Enemy

    forç_E = driver.find_element_by_xpath("/html/body/section[1]/div/div[2]/div[5]/ul[1]/li[2]/div/span[2]").text

    defe_E = driver.find_element_by_xpath("/html/body/section[1]/div/div[2]/div[5]/ul[1]/li[3]/div/span[2]").text

    agil_E = driver.find_element_by_xpath("/html/body/section[1]/div/div[2]/div[5]/ul[1]/li[4]/div/span[2]").text

    inte_E = driver.find_element_by_xpath("/html/body/section[1]/div/div[2]/div[5]/ul[1]/li[5]/div/span[2]").text

    resi_E = driver.find_element_by_xpath("/html/body/section[1]/div/div[2]/div[5]/ul[1]/li[6]/div/span[2]").text

    level_E = driver.find_element_by_xpath("/html/body/section[1]/div/div[2]/div[5]/div/div[1]/span[2]").text

    #nickname
    your_name = driver.find_element_by_xpath("/html/body/section[1]/div/div[2]/div[2]/div/div[2]").text
    enemy_name = driver.find_element_by_xpath("/html/body/section[1]/div/div[2]/div[5]/div/div[2]").text

    rows = "/html/body/section[2]/div/table/tbody/tr[{}]/td"
    
    
    try:
        for row in range (1,50):
            r_text = driver.find_element_by_xpath(rows.format(row)).text
            
            if r_text.split()[0] == nickname:
                r_text = r_text.replace(enemy_name,"")
                r_text = r_text.replace(your_name,"")
                damage = ''.join(x for x in r_text if x.isdigit())
                
                if damage != "":
                    sheet.write(row_A,1,damage)
                    sheet.write(row_A,2,forç)
                    sheet.write(row_A,3,defe)
                    sheet.write(row_A,4,agil)
                    sheet.write(row_A,5,inte)
                    sheet.write(row_A,6,resi)
                    sheet.write(row_A,7,level)
                    row_A += 1
            else:
                r_text = r_text.replace(enemy_name,"")
                r_text = r_text.replace(your_name,"")
                damage_e = ''.join(x for x in r_text if x.isdigit())
                
                if damage_e != "":
                    sheet.write(row_E,9,damage_e)
                    sheet.write(row_E,10,forç_E)
                    sheet.write(row_E,11,defe_E)
                    sheet.write(row_E,12,agil_E)
                    sheet.write(row_E,13,inte_E)
                    sheet.write(row_E,14,resi_E)
                    sheet.write(row_E,15,level_E)
                    row_E += 1
        
                    
    except:
        None
    wb.save('Dados.xls')  
    
    
#start
def run():   
    site_login(url)
    for i in range (1,3):
        get_page(i)
        combat_info()
    driver.close()
    os._exit(0)
    
    
    
    
run()
    
      

    
     
 

    
