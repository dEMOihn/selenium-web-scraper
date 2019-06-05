from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from openpyxl import Workbook
from constants import *

def main():
    driver = webdriver.Chrome(CHROM_DRIVER)
    driver.get(BASE_URL)

    list_of_countries = []
    list_of_comp = []
    list_of_services = []

    """Authorithation on the site"""
    driver.find_element_by_xpath("//*[@id='hs-eu-confirmation-button']").click()
    time.sleep(7)

    log_email = driver.find_element_by_name("email")
    log_email.send_keys(EMAIL)
    time.sleep(4)

    log_pass = driver.find_element_by_name("password")
    log_pass.send_keys(PASSWORD)
    log_pass.send_keys(Keys.ENTER)
    time.sleep(7)

    """Finding solutions on each page"""
    for page in range(1, 62):
        other_pages = BASE_URL + str(page)

        driver.get(other_pages)
        time.sleep(5)

        # Finding names of each company on the page and appending them to the list of companies
        companies = driver.find_elements_by_xpath("//div[@class='mb-2']//a")
        for c in companies:
            comp = c.text
            if "Sponsored" not in comp:
                list_of_comp.append(comp)

        # Finding names of each country on the page and appending them to the list of countries
        countries = driver.find_elements_by_xpath("//div[@class='left-bar-text d-block']//a")
        for i in countries:
            list_of_countries.append(i.text)

        # Finding services of all companies on the page and appending them to the list of services
        services = driver.find_elements_by_xpath("//div[@class='solution-about mb-3']")
        for service in services:
            list_of_services.append(service.text)

    driver.quit()

    """Writing base to the Excel document"""
    result = [list_of_countries, list_of_comp, list_of_services]
    i = 0
    columns = []
    while i <= len(list_of_countries)-1:
        columns.append([el[i] for el in result])
        i += 1

    wb = Workbook()
    ws = wb.active

    for subarray in columns:
        ws.append(subarray)

    wb.save("beesrrvice_base.xlsx")


if __name__=="__main__":
    main()

