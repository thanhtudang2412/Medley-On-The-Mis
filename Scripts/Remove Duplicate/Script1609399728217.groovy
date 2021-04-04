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
code = 7
duplicate = 21
fileName = 'D:\\Tú\\Processed\\LK_Aibiz_Tú_11.1.xlsx'
FileInputStream file = new FileInputStream(new File(fileName))
XSSFWorkbook workbook = new XSSFWorkbook(file)
XSSFSheet sheet = workbook.getSheetAt(0)
n = (sheet.getLastRowNum() + 1)
for (i = 1; i < n-1; i++) {
	now = sheet.getRow(i).getCell(code).getNumericCellValue()
	then = sheet.getRow(i+1).getCell(code).getNumericCellValue()
	if (now==then){
		sheet.getRow(i).createCell(duplicate).setCellValue(then)
		file.close()
		FileOutputStream outFile = new FileOutputStream(new File(fileName))
		workbook.write(outFile)
		outFile.close()
		file = new FileInputStream(new File(fileName))
		workbook = new XSSFWorkbook(file)
		sheet = workbook.getSheetAt(0)
	}
}
file.close()
FileOutputStream outFile = new FileOutputStream(new File(fileName))
workbook.write(outFile)
outFile.close()

