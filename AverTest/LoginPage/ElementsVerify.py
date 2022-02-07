from selenium import webdriver

driver=webdriver.Chrome(executable_path="C:/BIDS/beneficienttest/Beneficient/Chrome/chromedriver.exe")
driver.implicitly_wait(10)
driver.maximize_window()
driver.get("https://averreplica.1wayit.com/login")
check=driver.find_element_by_xpath("//img[@src='https://averreplica.1wayit.com/global_assets/images/logo.png']").is_displayed()
print(check)
driver.quit()
#driver.find_element_by_xpath("//input[@type='submit']").click()