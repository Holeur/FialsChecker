from selenium import webdriver
import bestdllever
import selenium
import openpyxl
import time
import os

browser = webdriver.Firefox()
browser.get('https://poe.ninja/challenge/vials')

wb = openpyxl.load_workbook('FialsTable.xlsx')
mainstr = wb.active

fials = [] #[['Название_фиала',цена_фиала],[]]

for num in range(1,10):                                          
    if browser.find_element_by_xpath("/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr["+str(num)+"]/td[3]/div/span[1]/img").get_attribute("title") == "Сфера хаоса":                                                                                                                                                                                        
        name = browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr['+str(num)+']/td[1]/div/div[1]').get_attribute('title')
        price = float(browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr['+str(num)+']/td[3]/div/span[1]').get_attribute('title'))
        fials.append([name,price])
        for num1 in range(1,23):
            if mainstr['A'+str(num1)].value == name:
                mainstr['B'+str(num1)].value = price
    else:         
        name = browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr['+str(num)+']/td[1]/div/div[1]').get_attribute('title')
        price = float(browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr['+str(num)+']/td[3]/div/span[2]').get_attribute('title'))
        fials.append([name,price])
        for num1 in range(1,23):
            if mainstr['A'+str(num1)].value == name:
                mainstr['B'+str(num1)].value = price

#print(fials)

def getItemNameByFial(fialname): #return ["Необходимый предмет для фиала","Результат обработки","Тип предмета"]
    if fialname == "Фиал призрака": 
        return "Душелов","Жнец душ","unique-flasks"
    elif fialname == "Фиал превосходства":
        return ["Закалённый дух","Закалённый разум","Закалённая плоть"],["Запредельный дух","Запредельный разум","Запредельная плоть"],"unique-jewels"
    elif fialname == "Фиал пробуждения":
        return "Забытьё Апепа","Главенство Апепа","unique-armours"
    elif fialname == "Фиал жертвоприношения":
        return "Жертвенное сердце","Сердце Зерфи","unique-accessories"
    elif fialname == "Фиал последствий":
        return "Оковы труса","Наследие труса","unique-accessories"
    elif fialname == "Фиал ритуала":
        return "Танец преподносимых","Омейокан","unique-armours"
    elif fialname == "Фиал призыва":
        return "Маска Духопийцы","Маска Сшитого Демона","unique-armours"
    elif fialname == "Фиал господства":
        return "Длань архитектора","Длань надзирателя","unique-armours"
    elif fialname == "Фиал судьбы":
        return "История ваал","Судьба ваал","unique-weapons"
    

    
#unique-flasks 
#unique-jewels
#unique-armours
#unique-accessories //*[@id="main"]/div/div[3]/div[1]/div[1]/input
#unique-weapons //*[@id="main"]/div/div[3]/div[1]/div[1]/input

resultprices = [] #["Имя фиала","Название предмета","Результат разницы"])

for fial in fials:
    #print('_______________________________________')
    #print(fial[0])
    namebaseitem = getItemNameByFial(fial[0])[0]
    #print(namebaseitem)
    nameresultitem = getItemNameByFial(fial[0])[1]
    #print(nameresultitem)
    typeitem = getItemNameByFial(fial[0])[2]
    #print(typeitem)

    browser.get('https://poe.ninja/challenge/'+typeitem)
    time.sleep(1)
    
    
    if fial[0] != "Фиал превосходства":
        browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/div/div[1]/input').clear()
        browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/div/div[1]/input').send_keys(namebaseitem)
        if browser.find_element_by_xpath("/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr/td[4]/div/span/img").get_attribute("title") == "Сфера хаоса":                                                                                                                                                                                        
            pricebaseitem = float(browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr/td[4]/div/span').get_attribute('title'))
        else:                                                                                                                                                 
            pricebaseitem = float(browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr/td[4]/div/span[2]').get_attribute('title'))
        for num1 in range(1,23):
            if mainstr['A'+str(num1)].value == namebaseitem:
                mainstr['B'+str(num1)].value = pricebaseitem
                
                
        browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/div/div[1]/input').clear()
        browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/div/div[1]/input').send_keys(nameresultitem)
        if browser.find_element_by_xpath("/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr/td[4]/div/span/img").get_attribute("title") == "Сфера хаоса":                                                                                                                                                                                        
            priceresultitem = float(browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr/td[4]/div/span').get_attribute('title'))
        else:                                                                                                                                  
            priceresultitem = float(browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr/td[4]/div/span[2]').get_attribute('title'))
        for num1 in range(1,23):
            if mainstr['E'+str(num1)].value == nameresultitem:
                mainstr['D'+str(num1)].value = priceresultitem
                
                
        baseandfial = fial[1] + pricebaseitem
        #print("Fial + Item = "+str(baseandfial))
        resultprice = priceresultitem - baseandfial
        #print("Result - (Fial + Item) = "+str(resultprice))
        resultprices.append([fial[0],namebaseitem,resultprice])
                
            
    else:
        for name in namebaseitem:
            browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/div/div[1]/input').clear()
            browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/div/div[1]/input').send_keys(name)
            if browser.find_element_by_xpath("/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr/td[3]/div/span/img").get_attribute("title") == "Сфера хаоса":                                                                                                                                                                                        
                pricebaseitem = float(browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr/td[3]/div/span').get_attribute('title'))
            else:                                                                                                                                                 
                pricebaseitem = float(browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr/td[3]/div/span[2]').get_attribute('title'))
                
            for num1 in range(1,23):
                if mainstr['A'+str(num1)].value == name:
                    mainstr['B'+str(num1)].value = pricebaseitem
                
                
            browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/div/div[1]/input').clear()
            browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/div/div[1]/input').send_keys(nameresultitem[namebaseitem.index(name)])
            if browser.find_element_by_xpath("/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr/td[3]/div/span/img").get_attribute("title") == "Сфера хаоса":                                                                                                                                                                                        
                priceresultitem = float(browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr/td[3]/div/span').get_attribute('title'))
            else:                                                                                                                                  
                priceresultitem = float(browser.find_element_by_xpath('/html/body/div[4]/div/div[2]/div[2]/table/tbody/tr/td[3]/div/span[2]').get_attribute('title'))
                
            for num1 in range(1,23):
                if mainstr['E'+str(num1)].value == nameresultitem[namebaseitem.index(name)]:
                    mainstr['D'+str(num1)].value = priceresultitem
                
                
            baseandfial = fial[1] + pricebaseitem
            #print("Fial + Item = "+str(baseandfial))
            resultprice = priceresultitem - baseandfial
            #print("Result - (Fial + Item) = "+str(resultprice))
            resultprices.append([fial[0],name,resultprice])
            

for result in resultprices:
    #print(result[2])
    mainstr["C"+str(-1+(resultprices.index(result)+1)*2)].value = result[2]
    
def save():
    try:
        wb.save('FialsTable.xlsx')
    except PermissionError:
        #print('Закройте таблицу')
        time.sleep(3)
        save()
    
save()
os.system('start FialsTable.xlsx')
browser.quit()
#print(resultprices)