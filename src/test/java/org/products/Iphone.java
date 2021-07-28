package org.products;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Iphone {
	public static void main(String[] args) throws IOException {
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\ELCOT\\eclipse-workspace\\Arjunan\\AmazonProducts\\driver\\chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		driver.get("http://www.Amazon.in/");
		WebElement searchBox = driver.findElement(By.id("twotabsearchtextbox"));
		searchBox.sendKeys("iphone",Keys.ENTER);
		List<WebElement> productNames = driver.findElements(By.xpath("//span[@class='a-size-medium a-color-base a-text-normal']"));
		File file=new File("C:\\Users\\ELCOT\\eclipse-workspace\\Arjunan\\AmazonProducts\\workbook\\mavenlaunch.xlsx");
		FileInputStream stream=new FileInputStream(file);
	    Workbook book=new XSSFWorkbook(stream);
	    Sheet createSheet = book.createSheet("productNames");
	    for (int i = 0; i < productNames.size(); i++) {
	    	WebElement webElement = productNames.get(i);
	    	String text = webElement.getText();
	    	Row row = createSheet.createRow(i);
	        Cell createCell = row.createCell(0);
	        createCell.setCellValue(text);
	        
		}
		FileOutputStream stream2=new FileOutputStream(file);
		book.write(stream2);
	    {
		

	}}
}