import datetime
import math
import re
import time
import openpyxl
from datetime import datetime,date
import datetime as datetime
from fpdf import FPDF
import pytest
from selenium import webdriver
import allure
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from sys import platform
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pyperclip
import random
import string


@allure.step("Entering username ")
def enter_username(username):
  driver.find_element_by_id("email").send_keys(username)

@allure.step("Entering password ")
def enter_password(password):
  driver.find_element_by_id("password").send_keys(password)

@pytest.fixture()
def test_setup():
  global driver
  global TestName
  global description
  global TestResult
  global TestResultStatus
  global TestDirectoryName
  global path

  TestName = "test_ElementsWorking"
  description = "This test scenario is to verify all the Working of Elements at Invoice Entry page"
  TestResult = []
  TestResultStatus = []
  TestFailStatus = []
  FailStatus="Pass"
  TestDirectoryName = "test_ElementsWorking"
  global Exe
  Exe="Yes"
  Directory = 'test_InvoiceEntry/'
  if platform == "linux" or platform == "linux2":
      path = '/home/legion/office 1wayit/AVER/AverTest/' + Directory
  elif platform == "win32" or platform == "win64":
      path = 'C:/AVER/AverTest/' + Directory

  ExcelFileName = "Execution"
  locx = (path+'Executiondir/' + ExcelFileName + '.xlsx')
  wbx = openpyxl.load_workbook(locx)
  sheetx = wbx.active

  for ix in range(1, 100):
      if sheetx.cell(ix, 1).value == None:
          break
      else:
          if sheetx.cell(ix, 1).value == TestName:
              if sheetx.cell(ix, 2).value == "No":
                  Exe="No"
              elif sheetx.cell(ix, 2).value == "Yes":
                  Exe="Yes"

  if Exe=="Yes":
      if platform == "linux" or platform == "linux2":
          driver = webdriver.Chrome(executable_path="/home/legion/office 1wayit/AVER/AverTest/chrome/chromedriverLinux1")
      elif platform == "win32" or platform == "win64":
          driver = webdriver.Chrome(executable_path="C:/AVER/AverTest/chrome/chromedriver.exe")

      driver.implicitly_wait(10)
      driver.maximize_window()
      driver.get("https://averreplica.1wayit.com/login")
      enter_username("admin@averplanning.com")
      enter_password("admin786")
      driver.find_element_by_xpath("//button[@type='submit']").click()

  yield
  if Exe == "Yes":
      ct = datetime.datetime.now().strftime("%d_%B_%Y_%I_%M%p")
      ctReportHeader = datetime.datetime.now().strftime("%d %B %Y %I %M%p")

      class PDF(FPDF):
          def header(self):
              self.image(path+'EmailReportContent/logo.png', 10, 8, 33)
              self.set_font('Arial', 'B', 15)
              self.cell(73)
              self.set_text_color(0, 0, 0)
              self.cell(35, 10, ' Test Report ', 1, 1, 'B')
              self.set_font('Arial', 'I', 10)
              self.cell(150)
              self.cell(30, 10, ctReportHeader, 0, 0, 'C')
              self.ln(20)

          def footer(self):
              self.set_y(-15)
              self.set_font('Arial', 'I', 8)
              self.set_text_color(0, 0, 0)
              self.cell(0, 10, 'Page ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')

      pdf = PDF()
      pdf.alias_nb_pages()
      pdf.add_page()
      pdf.set_font('Times', '', 12)
      pdf.cell(0, 10, "Test Case Name:  "+TestName, 0, 1)
      pdf.multi_cell(0, 10, "Description:  "+description, 0, 1)

      for i1 in range(len(TestResult)):
         pdf.set_fill_color(255, 255, 255)
         pdf.set_text_color(0, 0, 0)
         if (TestResultStatus[i1] == "Fail"):
             #print("Fill Red color")
             pdf.set_text_color(255, 0, 0)
             TestFailStatus.append("Fail")
         TestName1 = TestResult[i1].encode('latin-1', 'ignore').decode('latin-1')
         pdf.multi_cell(0, 7,str(i1+1)+")  "+TestName1, 0, 1,fill=True)
         TestFailStatus.append("Pass")
      pdf.output(TestName+"_" + ct + ".pdf", 'F')

      #-----------To check if any failed Test case present-------------------
      for io in range(len(TestResult)):
          if TestFailStatus[io]=="Fail":
              FailStatus="Fail"
      # ---------------------------------------------------------------------

      # -----------To add test case details in PDF details sheet-------------
      ExcelFileName = "FileName"
      loc = (path+'PDFFileNameData/' + ExcelFileName + '.xlsx')
      wb = openpyxl.load_workbook(loc)
      sheet = wb.active
      print()
      check = TestName
      PdfName = TestName + "_" + ct + ".pdf"
      checkcount = 0

      for i in range(1, 100):
          if sheet.cell(i, 1).value == None:
              if checkcount == 0:
                  sheet.cell(row=i, column=1).value = check
                  sheet.cell(row=i, column=2).value = PdfName
                  sheet.cell(row=i, column=3).value = TestDirectoryName
                  sheet.cell(row=i, column=4).value = description
                  sheet.cell(row=i, column=5).value = FailStatus
                  checkcount = 1
              wb.save(loc)
              break
          else:
              if sheet.cell(i, 1).value == check:
                  if checkcount == 0:
                    sheet.cell(row=i, column=2).value = PdfName
                    sheet.cell(row=i, column=3).value = TestDirectoryName
                    sheet.cell(row=i, column=4).value = description
                    sheet.cell(row=i, column=5).value = FailStatus
                    checkcount = 1
      #----------------------------------------------------------------------------

      #---------------------To add Test name in Execution sheet--------------------
      ExcelFileName1 = "Execution"
      loc1 = (path+'Executiondir/' + ExcelFileName1 + '.xlsx')
      wb1 = openpyxl.load_workbook(loc1)
      sheet1 = wb1.active
      checkcount1 = 0

      for ii1 in range(1, 100):
          if sheet1.cell(ii1, 1).value == None:
              if checkcount1 == 0:
                  sheet1.cell(row=ii1, column=1).value = check
                  checkcount1 = 1
              wb1.save(loc1)
              break
          else:
              if sheet1.cell(ii1, 1).value == check:
                  if checkcount1 == 0:
                    sheet1.cell(row=ii1, column=1).value = check
                    checkcount1 = 1
      #-----------------------------------------------------------------------------

      driver.quit()

@pytest.mark.smoke
def test_VerifyAllClickables(test_setup):
    if Exe == "Yes":
        TimeSpeed = 2
        SHORT_TIMEOUT = 3
        LONG_TIMEOUT = 60
        LOADING_ELEMENT_XPATH = "//div[@class='main-loader LoaderImageLogo']"
        try:
            # ---------------------------Verify Invoice Entry icon click-----------------------------
            PageName = "Invoice Entry icon"
            Ptitle1 = ""
            try:
                driver.find_element_by_xpath("//i[@class='icon-paragraph-justify3']/parent::a").click()
                time.sleep(2)
                driver.find_element_by_xpath("//div[@class='card card-sidebar-mobile']/ul/li[5]/a/i").click()
                time.sleep(2)
                driver.find_element_by_xpath("//div[@class='card card-sidebar-mobile']/ul/li[5]/ul/li[1]/a").click()
                time.sleep(2)
                TestResult.append(PageName + " is present in left menu and able to click")
                TestResultStatus.append("Pass")
            except Exception:
                TestResult.append(PageName + " is not present")
                TestResultStatus.append("Fail")
            print()
            time.sleep(TimeSpeed)
            # ---------------------------------------------------------------------------------

            # # ---------------------------Verify working of Back button on Invoice entry page -----------------------------
            # PageName = "Back button"
            # Ptitle1 = "Rae"
            # try:
            #     driver.find_element_by_xpath("//a[text()='Back']").click()
            #     time.sleep(2)
            #     PageTitle1 = driver.find_element_by_xpath("//div[@class='hed_wth_srch']/h2").text
            #     print(PageTitle1)
            #     assert PageTitle1 in Ptitle1, PageName + " not present"
            #     TestResult.append(PageName + " is clickable")
            #     TestResultStatus.append("Pass")
            # except Exception:
            #     TestResult.append(PageName + " is not clickable")
            #     TestResultStatus.append("Fail")
            # print()
            # time.sleep(TimeSpeed)
            # # ---------------------------------------------------------------------------------
            #
            # # ----------------Verify Invoice entry icon click after verifying back--------
            # PageName = "Invoice entry icon"
            # Ptitle1 = ""
            # try:
            #     driver.find_element_by_xpath("//i[@class='icon-paragraph-justify3']/parent::a").click()
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//div[@class='card card-sidebar-mobile']/ul/li[5]/a/i").click()
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//div[@class='card card-sidebar-mobile']/ul/li[5]/ul/li[1]/a").click()
            #     time.sleep(2)
            #     TestResult.append(PageName + "  is opened again after verifying back button")
            #     TestResultStatus.append("Pass")
            # except Exception:
            #     TestResult.append(PageName + " is not opened again after verifying back button")
            #     TestResultStatus.append("Fail")
            # print()
            # time.sleep(TimeSpeed)
            # # ---------------------------------------------------------------------------------
            #
            # # ---------------------------Verify working of XERO CSV button-----------------------------
            # PageName = "XERO CSV button"
            # Ptitle1 = "Generate Invoice Report"
            # try:
            #     driver.find_element_by_xpath("//button[text()='XERO CSV']").click()
            #     time.sleep(2)
            #     PageTitle1 = driver.find_element_by_xpath("//div[@id='xeroReport']/div/div/div[1]/h4").text
            #     time.sleep(2)
            #     print(PageTitle1)
            #     assert PageTitle1 in Ptitle1, PageName + " not present"
            #     TestResult.append(PageName + " is clickable")
            #     TestResultStatus.append("Pass")
            # except Exception:
            #     TestResult.append(PageName + " is not clickable")
            #     TestResultStatus.append("Fail")
            # print()
            # time.sleep(TimeSpeed)
            # CloseButton = driver.find_element_by_xpath("// div[ @ id = 'xeroReport'] / div / div / div[1] / button").click()
            # ---------------------------------------------------------------------------------

            # ---------------------------Verify working of Create new button on Invoice entry page-----------------------------
            PageName = "Create new button"
            Ptitle1 = "Create New Invoice"
            try:
                driver.find_element_by_xpath("//a[text()='Create New']").click()
                time.sleep(2)
                PageTitle1 = driver.find_element_by_xpath("//h2[text()='Create New Invoice']").text
                time.sleep(2)
                print(PageTitle1)

                assert PageTitle1 in Ptitle1, PageName + " not present"
                TestResult.append(PageName + " on Invoice entry page is clickable")
                TestResultStatus.append("Pass")
            except Exception:
                TestResult.append(PageName + " on Invoice entry page is not clickable")
                TestResultStatus.append("Fail")
            print()
            time.sleep(TimeSpeed)
            # ---------------------------------------------------------------------------------

            # # ---------------------------Verify working of Create Reimburse Client button on Create new page-----------------------------
            # PageName = "Create Reimburse Client button"
            # Ptitle1 = ""
            # try:
            #     driver.find_element_by_xpath("//a[@title='Create Reimburse Client']").click()
            #     time.sleep(2)
            #     for aa in range(5):
            #         letters = string.ascii_lowercase
            #         returna = ''.join(random.choice(letters) for i in range(5))
            #         EmailID = returna + "@gmail.com"
            #
            #         letters1 = string.ascii_lowercase
            #         returna1 = ''.join(random.choice(letters1) for i in range(5))
            #         Name = returna1
            #     driver.find_element_by_xpath("//input[@name='re_service_name']").send_keys(Name)
            #     time.sleep(2)
            #     print(PageTitle1)
            #     # select = Select(driver.find_element_by_xpath("//select[@name='re_service_status']"))
            #     # select.select_by_visible_text("Active")
            #     # time.sleep(2)
            #     SetUpDate = driver.find_element_by_xpath("//input[@name='re_set_up_data']")
            #     SetUpDate.clear()
            #     SetUpDate.send_keys("16-02-2022")
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='re_account_name']").send_keys(Name)
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='re_bsb']").send_keys("123456")
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='re_account_number']").send_keys("12345678")
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='re_send_remittance_email']").click()
            #     driver.find_element_by_xpath("//input[@name='re_send_remittance_email']").click()
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='re_remittance_email']").send_keys(EmailID)
            #     driver.find_element_by_xpath("//div[@id='createnewsplatest']/div/div/div[2]/form/div[last()]/button").click()
            #     assert PageTitle1 in Ptitle1, PageName + " not present"
            #     TestResult.append(PageName + "on create new invoice page is clickable and user able to create reimburse client ")
            #     TestResultStatus.append("Pass")
            # except Exception:
            #     TestResult.append(PageName + " is not clickable")
            #     TestResultStatus.append("Fail")
            # print()
            # time.sleep(TimeSpeed)
            # # ---------------------------------------------------------------------------------

            # # ---------------------------Verify working of Create Service Provider button on Create new page-----------------------------
            # PageName = "Create Service Provider button"
            # Ptitle1 = ""
            # try:
            #     driver.find_element_by_xpath("//a[@title='Create Service Provider']").click()
            #     time.sleep(2)
            #     for aa in range(5):
            #         letters = string.ascii_lowercase
            #         returna = ''.join(random.choice(letters) for i in range(5))
            #         EmailID = returna + "@gmail.com"
            #
            #         letters1 = string.ascii_lowercase
            #         returna1 = ''.join(random.choice(letters1) for i in range(5))
            #         Name = returna1
            #     driver.find_element_by_xpath("//input[@name='service_name']").send_keys(Name)
            #     time.sleep(2)
            #     print(PageTitle1)
            #     # select = Select(driver.find_element_by_xpath("//select[@name='re_service_status']"))
            #     # select.select_by_visible_text("Active")
            #     # time.sleep(2)
            #     SetUpDate = driver.find_element_by_xpath("//input[@name='set_up_data']")
            #     SetUpDate.clear()
            #     SetUpDate.send_keys("16-02-2022")
            #     # act = ActionChains(driver)
            #     # act.send_keys(Keys.ENTER).perform()
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='abn']").send_keys("123678")
            #     select = Select(driver.find_element_by_xpath("//select[@name='franchise']"))
            #     select.select_by_index(1)
            #     time.sleep(2)
            #     # select = Select(driver.find_element_by_xpath("//select[@name='related_franchise']"))
            #     # select.select_by_index(0)
            #     select = Select(driver.find_element_by_xpath("//select[@id='service_type']"))
            #     select.select_by_index(4)
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='office_number']").send_keys("1234567890")
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='mobile_number']").send_keys("1122334455")
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='service_email']").send_keys("test@abc.com")
            #     select = Select(driver.find_element_by_xpath("//select[@name='address_country']"))
            #     select.select_by_index(5)
            #     time.sleep(2)
            #     select = Select(driver.find_element_by_xpath("//select[@name='address_state']"))
            #     select.select_by_index(5)
            #     time.sleep(2)
            #     select = Select(driver.find_element_by_xpath("//select[@name='address_city']"))
            #     select.select_by_index(5)
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='address_street']").send_keys("XYZ")
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='postcode']").send_keys("225588")
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='suburb']").send_keys("147258")
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='account_name']").send_keys(Name)
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='bsb']").send_keys("123456")
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='account_number']").send_keys("12345678")
            #     time.sleep(2)
            #     CheckBox = driver.find_element_by_xpath("//input[@name='remittance_email']")
            #     CheckBox.clear()
            #     CheckBox.click()
            #     time.sleep(2)
            #     driver.find_element_by_xpath("//input[@name='remittance_email']").send_keys(EmailID)
            #     driver.find_element_by_xpath(
            #         "//div[@id='createnewsplatest']/div/div/div[2]/form/div[last()]/button").click()
            #     assert PageTitle1 in Ptitle1, PageName + " not present"
            #     TestResult.append(
            #         PageName + "on create new invoice page is clickable and user able to create reimburse client ")
            #     TestResultStatus.append("Pass")
            # except Exception:
            #     TestResult.append(PageName + " is not clickable")
            #     TestResultStatus.append("Fail")
            # print()
            # time.sleep(TimeSpeed)
            # # ---------------------------------------------------------------------------------

            # ---------------------------Verify working of Create New Invoice form-----------------------------
            PageName = "Create New Invoice form"
            Ptitle1 = ""
            l = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
            random.shuffle(l)
            if l[0] == 0:
                pos = random.choice(range(1, len(l)))
                l[0], l[pos] = l[pos], l[0]
            InvoiceNumber = ''.join(map(str, l[0:4]))
            print(InvoiceNumber)
            try:
                act = ActionChains(driver)
                driver.find_element_by_xpath("//input[@name='search_client_name']").send_keys("SUNIL")
                time.sleep(2)
                act.send_keys(Keys.DOWN).perform()
                act.send_keys(Keys.ENTER).perform()
                time.sleep(2)
                driver.find_element_by_xpath("//input[@name='provider_invoice_number']").send_keys(InvoiceNumber)
                time.sleep(2)
                driver.find_element_by_xpath("//input[@name='search_service_provider_name']").send_keys("Blossom")
                time.sleep(2)
                act.send_keys(Keys.DOWN).perform()
                act.send_keys(Keys.ENTER).perform()
                time.sleep(2)
                today = date.today()
                D1 = today.strftime("%d-%m-%Y")
                driver.find_element_by_xpath("//input[@name='provider_invoice_date']").send_keys(D1)
                time.sleep(2)
                act.send_keys(Keys.ENTER).perform()
                select = Select(driver.find_element_by_xpath("//select[@id='approvedByClientText']"))
                select.select_by_index(0)
                time.sleep(2)
                driver.find_element_by_xpath("//div[@id='search_client_data']/div[2]/div/div[1]/h3").click()
                time.sleep(2)
                driver.find_element_by_xpath("//div[@id='search_client_data']/div[2]/div/div[1]/div/textarea").send_keys("Test Notes")
                time.sleep(2)
                driver.find_element_by_xpath("//div[@id='search_client_data']/div[2]/div/div[3]/h3").click()
                time.sleep(2)
                driver.find_element_by_xpath("//div[@id='search_client_data']/div[2]/div/div[3]/div/div/div[1]/label/input").click()
                time.sleep(2)
                driver.find_element_by_xpath("//input[@name='service_detail[2][delivered_date]']").send_keys("05-02-2022")
                time.sleep(2)
                driver.find_element_by_xpath("//span[@title='Select Category']").click()
                time.sleep(2)
                driver.find_element_by_xpath("//input[@class='select2-search__field']").send_keys("Transport")
                time.sleep(2)
                act.send_keys(Keys.ENTER).perform()
                driver.find_element_by_xpath("//span[@title='Select Item Number']").click()
                time.sleep(2)
                act.send_keys(Keys.DOWN).perform()
                act.send_keys(Keys.DOWN).perform()
                time.sleep(2)
                act.send_keys(Keys.ENTER).perform()
                driver.find_element_by_xpath("//input[@name='service_detail[2][qty]']").send_keys("1")
                time.sleep(2)
                driver.find_element_by_xpath("//input[@name='service_detail[2][price]']").send_keys("5")
                time.sleep(2)
                assert PageTitle1 in Ptitle1, PageName + " not present"
                TestResult.append(PageName + " on Invoice entry page is working")
                TestResultStatus.append("Pass")
            except Exception:
                TestResult.append(PageName + " on Invoice entry page is not working")
                TestResultStatus.append("Fail")
            print()
            time.sleep(TimeSpeed)
            # ---------------------------------------------------------------------------------

        except Exception as err:
            print(err)
            TestResult.append("Invoice entry is not working correctly. Below error found\n"+str(err))
            TestResultStatus.append("Fail")
            pass

    else:
        print()
        print("Test Case skipped as per the Execution sheet")
        skip = "Yes"

        # -----------To add Skipped test case details in PDF details sheet-------------
        ExcelFileName = "FileName"
        loc = (path+'PDFFileNameData/' + ExcelFileName + '.xlsx')
        wb = openpyxl.load_workbook(loc)
        sheet = wb.active
        check = TestName

        for i in range(1, 100):
            if sheet.cell(i, 1).value == check:
                sheet.cell(row=i, column=5).value = "Skipped"
                wb.save(loc)
        # ----------------------------------------------------------------------------


