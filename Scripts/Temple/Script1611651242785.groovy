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

WebUI.openBrowser('')

WebUI.navigateToUrl('http://203.169.117.254:8080/login?forcedLogin=true')

WebUI.setText(findTestObject('Temple/Page_MISASIA VER2.0/input_Login_username'), 'dangthanhtu')

WebUI.setEncryptedText(findTestObject('Temple/Page_MISASIA VER2.0/input_Login_password'), 'dR8VJiTWKXduFNTM4d68fg==')

WebUI.click(findTestObject('Temple/Page_MISASIA VER2.0/input_Remember me_Submit'))

WebUI.delay(15)
 
WebUI.click(findTestObject('Temple/Page_MISASIA  Admin/a_Application'))


WebUI.click(findTestObject('Temple/Page_MISASIA  Admin/a_Work Retrieval'))

WebUI.delay(10)

WebUI.setText(findTestObject('Temple/Page_MIS  WORK RETRIEVAL/input_Works Internal No._searchWkIntNoFilterVal'), 
    '19860601')

WebUI.click(findTestObject('Temple/Page_MIS  WORK RETRIEVAL/input_Settings_btn btn-info pull-left searchButton'))

a = WebUI.getAttribute(findTestObject('Page_MIS  WORK EDIT-19860601/input_Title_workOTTitleEng validatecustomcu_eb42fe'), '')

println(a)

