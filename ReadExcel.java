package ReadExcelData;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;


public class ReadExcel {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File src = new File("E:\\Excel.xlsx");
		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		Scanner scan = new Scanner(System.in);
		System.out.println("Enter today name:");
		String date = scan.next();
		XSSFSheet sheet1 = wb.getSheet(date);
		for (int i = 0; i < 10; i++) {
			String data0 = sheet1.getRow(2 + i).getCell(2).getStringCellValue();
			System.out.println("Keyword: " + data0);

			System.setProperty("webdriver.chrome.driver",
					"E:\\selenium webdriver\\chromedriver\\chromedriver_win32\\chromedriver.exe");
			WebDriverManager.chromedriver();
			WebDriver driver = new ChromeDriver();
			driver.get("https://www.google.com");
			
			Thread.sleep(1000);
			driver.findElement(By.xpath("//input")).sendKeys(data0);
			//driver.findElement(By.xpath("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input"))
					//.sendKeys(data0);
			Thread.sleep(1000);
			List<WebElement> elements = driver.findElements(
					By.xpath("//ul[@role='listbox']//li/descendant::div[@class='wM6W7d']"));
			
			System.out.println(elements.size());
			WebElement q;
			
			for (int j = 0; j < elements.size(); j++) {
				q = elements.get(j);
				
				
				System.out.println(q.getText());
			}
		int sm=0;int b=0; int flags=0; int flagb=0;
			for (int j = 0; j < elements.size(); j++) {
				q = elements.get(j);
				String s=q.getText();
				int n=s.length();
				//System.out.println("length "+n);
				if(sm<=0) {sm=n; flags=j;}
				else if(sm>n) {sm=n; flags=j; }
				if(b<n) {b=n; flagb=j;}
			}
			
			q = elements.get(flags);
			
			System.out.println("Shortest option: "+ q.getText());
			sheet1.getRow(2 + i).createCell(4).setCellValue(q.getText());
			FileOutputStream fos = new FileOutputStream(src);
			wb.write(fos);
			q = elements.get(flagb);
			
			System.out.println("Longest option: "+ q.getText());
			sheet1.getRow(2 + i).createCell(3).setCellValue(q.getText());
			FileOutputStream foos = new FileOutputStream(src);
			wb.write(foos);
			driver.close();
		}	
	}

}
