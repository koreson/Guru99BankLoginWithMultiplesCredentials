package utils;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelDataProvider {

	WebDriver driver;

	@BeforeTest
	public void setUpTest() {
		String projectPath = System.getProperty("user.dir");
		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\User\\Downloads\\chromedriver_win32\\chromedriver.exe");
		driver = new ChromeDriver();
	}

	@Test(dataProvider = "test1data")
	public void test1(String userid, String password) throws Exception {
		System.out.println(userid + "  |  " + password);

		driver = new ChromeDriver();
		driver.manage().deleteAllCookies();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);

		driver.get("http://www.demo.guru99.com/V4/");

		driver.findElement(By.xpath("//input[@name='uid']")).sendKeys(userid);

		driver.findElement(By.xpath("//input[@name='password']")).sendKeys(password);    

		driver.findElement(By.xpath("//input[@name='btnLogin']")).click();  

		//driver.quit();
	}

	@DataProvider(name = "test1data")
	public Object[][] getData() throws IOException {
		String excelPath = "C:\\Users\\User\\workspace\\Guru99BankLoginWithMultiplesCredentials\\excelgurubank99\\testdata.xlsx";
		Object data[][] = testData(excelPath, "Sheet1");
		return data;
	}

	public Object[][] testData(String excelPath, String SheetName) throws IOException {

		ExcelUtils excel = new ExcelUtils(excelPath, SheetName);

		int rowCount = excel.getRowCount();
		int colCount = excel.getColCount();

		Object data[][] = new Object[rowCount - 1][colCount];

		for (int i = 1; i < rowCount; i++) {
			for (int j = 0; j < colCount; j++) {

				String cellData = excel.getCellDataString(i, j);
				System.out.print(cellData + " | ");
				data[i - 1][j] = cellData;
			}

			// System.out.println();
		}
		return data;

	}
}
