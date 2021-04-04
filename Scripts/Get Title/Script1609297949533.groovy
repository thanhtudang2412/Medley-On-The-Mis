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

import java.text.Normalizer;
import java.util.regex.Pattern;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.lang.String

public class NlpUtils{
	public static String removeAccent(String s) { String temp = Normalizer.normalize(s, Normalizer.Form.NFD); Pattern pattern = Pattern.compile("\\p{InCombiningDiacriticalMarks}+"); temp = pattern.matcher(temp).replaceAll("");
		temp = temp.replaceAll("đ", "d");
		return temp.replaceAll("Đ", "D"); }
	
	 public static void main(String []args){
	 }
}

WebUI.openBrowser('')
fileName = "D:\\Tú\\Processed\\LK_Aibiz_Tú_18.1.xlsx"
idLink = 3
vieTitle = 17
engTitle = 18
FileInputStream file = new FileInputStream (new File(fileName))
XSSFWorkbook workbook = new XSSFWorkbook(file);
XSSFSheet sheet = workbook.getSheetAt(0);
n = sheet.getLastRowNum()+1
println(n)
NlpUtils m = new NlpUtils()
preRowVal = "pre"
rowVal = "start"
for (i = 1; i<n; i++){
	rowVal = sheet.getRow(i).getCell(idLink).getStringCellValue();
	if (rowVal!=preRowVal){
		WebUI.navigateToUrl("https://www.youtube.com/watch?v=" + rowVal)
		s = WebUI.getWindowTitle()
		String [] a = s.split(' - You')
		s = a[0]
		s = s.replaceAll("[^\\p{L}\\d\\s]", " ");
		s = s.replaceAll(" +"," ");
		s = s.replaceAll("Liên Khúc","LK");
		s = s.replaceAll("Liên khúc","LK");
		s = s.replaceAll("liên khúc","LK");
		s = s.replaceAll("Lk","LK");
		if (s.contains("LK")){
			s = s.substring(s.lastIndexOf("LK"))
		}
		else{
			s = "LK " + s
		}
		sheet.getRow(i).createCell(vieTitle).setCellValue(s);
		s = m.removeAccent(s)
		sheet.getRow(i).createCell(engTitle).setCellValue(s+' ('+rowVal+') ');
		preRowVal = rowVal
		
		file.close();
		FileOutputStream outFile = new FileOutputStream(new File(fileName));
		workbook.write(outFile);
		outFile.close();
		
		file = new FileInputStream (new File(fileName))
		workbook = new XSSFWorkbook(file);
		sheet = workbook.getSheetAt(0);
	}
}
WebUI.closeBrowser()
file.close();
FileOutputStream outFile = new FileOutputStream(new File(fileName));
workbook.write(outFile);
outFile.close();