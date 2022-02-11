import datetime
import math
import re
import time
import openpyxl
from fpdf import FPDF
import pytest
from selenium import webdriver
import allure
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


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
  description = "This test scenario is to verify all the Working of Elements at Communication log of application"
  TestResult = []
  TestResultStatus = []
  TestFailStatus = []
  FailStatus="Pass"
  TestDirectoryName = "test_ElementsWorking"
  global Exe
  Exe="Yes"
  Directory = 'test_CommunicationLog/'
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
      driver=webdriver.Chrome(executable_path="C:/AVER/AverTest/chrome/chromedriver.exe")
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

      #driver.quit()

@pytest.mark.smoke
def test_VerifyAllClickables(test_setup):
    if Exe == "Yes":
        SHORT_TIMEOUT = 2
        LONG_TIMEOUT = 60
        LOADING_ELEMENT_XPATH = "//div[@class='main-loader LoaderImageLogo']"
        try:
            # ---------------------------Verify Communication Log icon click-----------------------------
            PageName = "Communication log icon"
            Ptitle1 = ""
            try:
                driver.find_element_by_xpath("//div[@class='card card-sidebar-mobile']/ul/li[2]/a/i").click()
                time.sleep(2)
                TestResult.append(PageName + "  is present in left menu and able to click")
                TestResultStatus.append("Pass")
            except Exception:
                TestResult.append(PageName + " is not present")
                TestResultStatus.append("Fail")
            print()
            # ---------------------------------------------------------------------------------
            # # ---------------------------Verify Working of Back button on Communication log page-----------------------------
            # PageName = "Back button on Communication log page"
            # Ptitle1 = "Rae"
            # try:
            #     driver.find_element_by_xpath("//div[@class='content yellow_color']/div/div[2]/div/a").click()
            #     time.sleep(2)
            #     PageTitle1 = driver.find_element_by_xpath("//div[@class='hed_wth_srch']/h2").text
            #     print(PageTitle1)
            #     assert PageTitle1 in Ptitle1, PageName + " not "
            #     TestResult.append(PageName + "  is clickable")
            #     TestResultStatus.append("Pass")
            # except Exception:
            #     TestResult.append(PageName + " is not clickable")
            #     TestResultStatus.append("Fail")
            # print()
            # # ---------------------------------------------------------------------------------
            #
            # # ---------------------------Verify Communication Log icon click after verifying back-----------------------------
            # PageName = "Communication log icon"
            # Ptitle1 = ""
            # try:
            #     driver.find_element_by_xpath("//div[@class='card card-sidebar-mobile']/ul/li[2]/a/i").click()
            #     time.sleep(2)
            #     TestResult.append(PageName + "  is opened again after verifying back button")
            #     TestResultStatus.append("Pass")
            # except Exception:
            #     TestResult.append(PageName + " is not opened again after verifying back button")
            #     TestResultStatus.append("Fail")
            # print()
            # # ---------------------------------------------------------------------------------
            # # ---------------------------Verify Select entry dropdown working-----------------------------
            # for cv in range (5):
            #     DropdownValues = {"Email": "//select[@name='communication_type']/option","Letter": "//select[@name='communication_type']/option","SMS": "//select[@name='communication_type']/option    ","Phone Call": "//select[@name='communication_type']/option","New Plan Form": "//ul[@class='GeneralClientDetails']/li[1]"}
            #     select = Select(driver.find_element_by_xpath(
            #         "//div[@class='content yellow_color']/div[1]/div/form/div/div/div/select"))
            #     Selector=['Email','Letter','SMS','Phone Call','New Plan Form']
            #     if cv==0:
            #         select.select_by_visible_text(Selector[3])
            #         TextCheck=Selector[3]
            #         path=DropdownValues[Selector[3]]
            #     elif cv==1:
            #         select.select_by_visible_text(Selector[0])
            #         TextCheck = Selector[0]
            #         path =DropdownValues[Selector[0]]
            #     elif cv==2:
            #         select.select_by_visible_text(Selector[1])
            #         TextCheck = Selector[1]
            #         path =DropdownValues[Selector[1]]
            #     elif cv==3:
            #         select.select_by_visible_text(Selector[2])
            #         TextCheck = Selector[2]
            #         path =DropdownValues[Selector[2]]
            #     elif cv==4:
            #         select.select_by_visible_text(Selector[4])
            #         TextCheck = Selector[4]
            #         path =DropdownValues[Selector[4]]
            #
            #
            #     textFound=driver.find_element_by_xpath(path).text
            #     if ":" in textFound:
            #         textFound=textFound.split(":")
            #         textFound=textFound[1]
            #
            #     textFound=textFound.strip()
            #     print(textFound)
            #     if textFound==TextCheck:
            #         print("aaaaa")
            #         TestResult.append(TextCheck + " dropdown value inside --- is able to click")
            #         TestResultStatus.append("Pass")
            #     else:
            #         print("ccccc")
            #         TestResult.append(TextCheck + " dropdown value inside --- is not able to click and open")
            #         TestResultStatus.append("Fail")
            #     driver.find_element_by_xpath("//div[@class='card card-sidebar-mobile']/ul/li[2]/a/i").click()
            #     time.sleep(2)
            # print()
            # # ---------------------------------------------------------------------------------
            # ---------------------------Verify Choose template dropdown working-----------------------------
            for dct in range(2):
                DropdownCT = {"General": "//div[@class='content yellow_color']/div[2]/div[1]/div/div/h2/small",
                                  "New Plan Form": "//div[@class='content yellow_color']/div[2]/div[1]/div/div/h2/small"}
                select = Select(driver.find_element_by_xpath(
                    "//div[@class='content yellow_color']/div[2]/div/div/div/form/div/div/select"))
                Selector = ['General', 'New Plan Form']
                if dct == 0:
                    select.select_by_visible_text(Selector[1])
                    TextCheck = Selector[1]
                    path = DropdownCT[Selector[1]]
                elif dct == 1:
                    select.select_by_visible_text(Selector[0])
                    TextCheck = Selector[0]
                    path = DropdownCT[Selector[0]]
                textFound = driver.find_element_by_xpath(path).text

                print(textFound)
                if textFound == TextCheck:
                    print("aaaaa")
                    TestResult.append(TextCheck + " dropdown value inside --- is able to click")
                    TestResultStatus.append("Pass")
                else:
                    print("ccccc")
                    TestResult.append(TextCheck + " dropdown value inside --- is not able to click and open")
                    TestResultStatus.append("Fail")
                driver.find_element_by_xpath("//div[@class='card card-sidebar-mobile']/ul/li[2]/a/i").click()
                time.sleep(2)
            print()
            # ---------------------------------------------------------------------------------



        except Exception as err:
            print(err)
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


