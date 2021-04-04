import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import org.apache.poi.xssf.usermodel.XSSFCell as XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow as XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet as XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook as XSSFWorkbook
import org.junit.After as After
import java.lang.String as String
import java.text.Normalizer as Normalizer
import java.util.regex.Pattern as Pattern
import java.io.FileInputStream as FileInputStream
import java.io.FileNotFoundException as FileNotFoundException
import java.io.IOException as IOException
import java.util.Date as Date
import com.kms.katalon.core.testobject.ConditionType as ConditionType

WebUI.openBrowser('')

WebUI.navigateToUrl('http://203.169.117.254:8080/login?forcedLogin=true')

WebUI.setText(findTestObject('Object Repository/Edit Medley/Page_MISASIA VER2.0/input_Login_username'), 'dangthanhtu')

WebUI.setEncryptedText(findTestObject('Object Repository/Edit Medley/Page_MISASIA VER2.0/input_Login_password'), 'dR8VJiTWKXduFNTM4d68fg==')

WebUI.click(findTestObject('Object Repository/Edit Medley/Page_MISASIA VER2.0/input_Remember me_Submit'))

WebUI.click(findTestObject('Object Repository/Edit Medley/Page_MISASIA  Admin/a_Application'))

WebUI.click(findTestObject('Object Repository/Edit Medley/Page_MISASIA  Admin/a_Work Retrieval'))

WebUI.delay(10)

fileName = 'D:\\Tú\\Processed\\WKS_COMPOSITE4.xlsx'

medleyCode = 0

engTitleNum = 1

vieTitleNum = 2

done = 7

FileInputStream file = new FileInputStream(new File(fileName))

XSSFWorkbook workbook = new XSSFWorkbook(file)

XSSFSheet sheet = workbook.getSheetAt(0)

n = (sheet.getLastRowNum() + 1)

for (i = 158; i < n; i++) {
    code = sheet.getRow(i).getCell(medleyCode).getNumericCellValue().toInteger()

    //   try {
    WebUI.setText(findTestObject('Object Repository/Edit Medley/Page_MIS  WORK RETRIEVAL/input_Works Internal No._searchWkIntNoFilterVal'), 
        code.toString())

    WebUI.click(findTestObject('Object Repository/Edit Medley/Page_MIS  WORK RETRIEVAL/input_Settings_btn btn-info pull-left searchButton'))

    WebUI.delay(12)

	engTitle = sheet.getRow(i).getCell(engTitleNum).getStringCellValue()
//	vieTitle = sheet.getRow(i).getCell(vieTitleNum).getStringCellValue()
    WebUI.setText(findTestObject('Object Repository/Edit Medley/Page_MIS  WORK EDIT-19822034/input_Title_workOTTitleEng validatecustomcu_eb42fe'), engTitle)
//
//    String[] a = engTitle.split(' ')
//
//    String[] b = vieTitle.split(' ')
//	
//	vieTitle = ""
//	
//    for (j = 0; j < a.length; j++) {
//        vieTitle += (b[j]+" ")
//    }
//	vieTitle = vieTitle.trim() 
//
//	WebUI.setText(findTestObject('Object Repository/Edit Medley/Page_MIS  WORK EDIT-19822034/input_Type of Work_workOTTitleLocal chinese_da3702'), vieTitle)

    WebUI.click(findTestObject('Object Repository/Edit Medley/Page_MIS  WORK EDIT-19822034/li_COMMIT'))

	WebUI.delay(20)
	
    WebUI.click(findTestObject('Object Repository/Edit Medley/Page_MIS  WORK EDIT-19822034/span_WORK EDIT-198..._ui-icon ui-icon-circle-close'))

    sheet.getRow(i).createCell(done).setCellValue('Done')

    file.close()

    FileOutputStream outFile = new FileOutputStream(new File(fileName))

    workbook.write(outFile)

    outFile.close()

    file = new FileInputStream(new File(fileName))

    workbook = new XSSFWorkbook(file)

    sheet = workbook.getSheetAt(0)
/*   }
       catch (Exception e) {
           continue
       } */
}

WebUI.closeBrowser()

file.close()

FileOutputStream outFile = new FileOutputStream(new File(fileName))

workbook.write(outFile)

outFile.close()

