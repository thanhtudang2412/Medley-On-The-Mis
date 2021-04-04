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
import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import java.text.Normalizer as Normalizer
import java.util.regex.Pattern as Pattern
import java.io.FileInputStream as FileInputStream
import java.io.FileNotFoundException as FileNotFoundException
import java.io.IOException as IOException
import java.util.Date as Date
import org.apache.poi.xssf.usermodel.XSSFCell as XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow as XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet as XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook as XSSFWorkbook
import java.lang.String as String

WebUI.openBrowser('')

WebUI.navigateToUrl('http://203.169.117.254:8080/login?forcedLogin=true')

WebUI.setText(findTestObject('Object Repository/Make Medley/Page_MISASIA VER2.0/input_Login_username'), 'dangthanhtu')

WebUI.setEncryptedText(findTestObject('Object Repository/Make Medley/Page_MISASIA VER2.0/input_Login_password'), 'dR8VJiTWKXduFNTM4d68fg==')

WebUI.click(findTestObject('Object Repository/Make Medley/Page_MISASIA VER2.0/input_Remember me_Submit'))

WebUI.delay(2)

WebUI.click(findTestObject('Object Repository/Make Medley/Page_MISASIA  Admin/a_Application'))

WebUI.click(findTestObject('Object Repository/Make Medley/Page_MISASIA  Admin/a_Work Retrieval'))

fileName = "D:\\Tú\\Processed\\LK_Aibiz_Tú_18.1.xlsx"
idLink = 3
vieTitle = 17
engTitle = 18
code = 6
failedCode = 19
addDelay = 0.75
createDelay = 10
commitDelay = 20 
FileInputStream file = new FileInputStream(new File(fileName))
XSSFWorkbook workbook = new XSSFWorkbook(file)
XSSFSheet sheet = workbook.getSheetAt(0)
n = (sheet.getLastRowNum() + 1)
println(n)

preRowVal = 'pre'

rowVal = 'start'

for (i = 1; i < n; i++) {
	rowVal = sheet.getRow(i).getCell(idLink).getStringCellValue()

	if (rowVal != preRowVal) {
		if (preRowVal != 'pre') {
			WebUI.click(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/li_COMMIT'))
			WebUI.delay(commitDelay)
			WebUI.click(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/span_WORK ED..._ui-icon ui-icon-circle-close'))
			WebUI.delay(1)
		}
		
		WebUI.click(findTestObject('Object Repository/Make Medley/Page_MIS  WORK RETRIEVAL/li_CREATE NEW'))
		WebUI.delay(createDelay)
		
		WebUI.setText(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/input_Title_workOTTitleEng validatecustomcu_eb42fe'),
			sheet.getRow(i).getCell(engTitle).getStringCellValue())
		
		WebUI.setText(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/input_Type of Work_workOTTitleLocal chinese_da3702'),
			sheet.getRow(i).getCell(vieTitle).getStringCellValue())
		
		WebUI.selectOptionByValue(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/select_SELECTPOPULARSERIOUSJAZZUNCLASSIFIED'),
			'UNC', true)
		
		WebUI.delay(1)
		
		WebUI.selectOptionByValue(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/select_SELECTACTIONACTION DRAMAACTION SUSPE_ad90cb'),
			'PP', true)
		
		WebUI.selectOptionByValue(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/select_SELECT(AFAN) OROMOABKHAZIANAFARAFRIK_5fa68f'),
			'VI', true)
		
		WebUI.selectOptionByValue(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/select_SELECT(AFAN) OROMOABKHAZIANAFARAFRIK_5fa68f_1'),
			'VI', true)
		
		WebUI.click(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/input_Yes_composite'))
		
		WebUI.selectOptionByValue(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/select_SELECTComposite of SampleMedleyPot-P_fc854d'),
			'MED', true)
		
		WebUI.setText(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/textarea_Remarks_workRemarks width60pc'),
			sheet.getRow(i).getCell(idLink).getStringCellValue())
		
		WebUI.setText(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/input_Internal No._componentWorkIntNo'),
			'100000')
		
		WebUI.delay(1)
		
		preRowVal = rowVal
		
		try{
			WebUI.setText(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/input_Internal No._componentWorkIntNo'),
				sheet.getRow(i).getCell(code).getNumericCellValue().intValue().toString())
			
			WebUI.delay(addDelay)
			
			WebUI.click(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/a_Ttl (Num, Type)  (1, OT)  100000-LARENZO (THEME)'))
			
			WebUI.delay(0.1)
			
			WebUI.click(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/img_Add Component_width80pc changeColorOfAddIcon'))
	
			WebUI.delay(0.1)
			
		}catch (Exception e) {
				sheet.getRow(i).createCell(failedCode).setCellValue(code)
                continue
        } 	
	}else{
		try{
			WebUI.setText(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/input_Internal No._componentWorkIntNo'),
				sheet.getRow(i).getCell(code).getNumericCellValue().intValue().toString())
		
			WebUI.delay(addDelay)
		
			WebUI.click(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/a_Ttl (Num, Type)  (1, OT)  100000-LARENZO (THEME)'))
		
			WebUI.delay(0.1)
		
			WebUI.click(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/img_Add Component_width80pc changeColorOfAddIcon'))

			WebUI.delay(0.1)
		
		}catch (Exception e) {
			sheet.getRow(i).createCell(failedCode).setCellValue(code)
			file.close()
			FileOutputStream outFile = new FileOutputStream(new File(fileName))
			workbook.write(outFile)
			outFile.close()
			file = new FileInputStream(new File(fileName))
			workbook = new XSSFWorkbook(file)
			sheet = workbook.getSheetAt(0)
			continue
		}
	}
}

WebUI.click(findTestObject('Object Repository/Make Medley/Page_MIS  WORK CREATE/li_COMMIT'))
WebUI.delay(20)
WebUI.closeBrowser()
file.close()
FileOutputStream outFile = new FileOutputStream(new File(fileName))
workbook.write(outFile)
outFile.close()







