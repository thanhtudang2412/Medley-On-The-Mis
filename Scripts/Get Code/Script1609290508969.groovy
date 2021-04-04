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
import java.lang.String as String
import java.text.Normalizer as Normalizer
import java.util.regex.Pattern as Pattern
import java.io.FileInputStream as FileInputStream
import java.io.FileNotFoundException as FileNotFoundException
import java.io.IOException as IOException
import java.util.Date as Date

WebUI.openBrowser('')

WebUI.navigateToUrl('http://203.169.117.254:8080/login?forcedLogin=true')

WebUI.setText(findTestObject('Object Repository/Get Code/Page_MISASIA VER2.0/input_Login_username'), 'dangthanhtu')

WebUI.setEncryptedText(findTestObject('Object Repository/Get Code/Page_MISASIA VER2.0/input_Login_password'), 'dR8VJiTWKXduFNTM4d68fg==')

WebUI.click(findTestObject('Object Repository/Get Code/Page_MISASIA VER2.0/input_Remember me_Submit'))

WebUI.delay(5)

WebUI.click(findTestObject('Object Repository/Get Code/Page_MISASIA  Admin/a_Application'))


WebUI.click(findTestObject('Object Repository/Get Code/Page_MISASIA  Admin/a_Work Retrieval'))

WebUI.delay(5)

fileName = "D:\\Tú\\Processed\\LK_Aibiz_Tú_18.1.xlsx"
medleyCode = 16
idLink = 3
vieTitle = 17
engTitle = 18
FileInputStream file = new FileInputStream(new File(fileName))

XSSFWorkbook workbook = new XSSFWorkbook(file)

XSSFSheet sheet = workbook.getSheetAt(0)

n = (sheet.getLastRowNum() + 1)

println(n)

preRowVal = 'pre'

rowVal = 'start'

for (i = 1; i < n; i++) {
	rowVal = sheet.getRow(i).getCell(idLink).getStringCellValue()
		try {
			WebUI.setText(findTestObject('Object Repository/Get Code/Page_MIS  WORK RETRIEVAL/input_Title_searchTitleFilterVal'), sheet.getRow(i).getCell(engTitle).getStringCellValue())
			
			WebUI.click(findTestObject('Object Repository/Get Code/Page_MIS  WORK RETRIEVAL/input_Settings_btn btn-info pull-left searchButton'))
			
			WebUI.delay(1)

			s = WebUI.getWindowTitle()

			String[] a = s.split('-')

			s = (a[1])

			sheet.getRow(i).createCell(medleyCode).setCellValue(s)

			// WebUI.delay(1)

			WebUI.click(findTestObject('Object Repository/Get Code/Page_MIS  WORK EDIT-20164182/span_WORK EDIT-201..._ui-icon ui-icon-circle-close'))
			
			file.close()
			
			FileOutputStream outFile = new FileOutputStream(new File(fileName))
			
			workbook.write(outFile)
			
			outFile.close()
			
			file = new FileInputStream(new File(fileName))
			
			workbook = new XSSFWorkbook(file)
			
			sheet = workbook.getSheetAt(0)
			
		}
		catch (Exception e) {
			continue
		}
		
	
}

WebUI.closeBrowser()

file.close()

FileOutputStream outFile = new FileOutputStream(new File(fileName))

workbook.write(outFile)

outFile.close()









