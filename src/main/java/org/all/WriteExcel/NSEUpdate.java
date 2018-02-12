package org.all.WriteExcel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class NSEUpdate {
	public static void main(String[] args) throws IOException {
System.setProperty("webdriver.chrome.driver",
			"C:/Users/Aparna/eclipse-workspace/TrialFB/Drivers/chromedriver.exe");
 WebDriver driver = new ChromeDriver();
driver.get("https://www.nseindia.com/live_market/dynaContent/live_watch/equities_stock_watch.htm?cat=N"); 
Workbook w=new XSSFWorkbook();
Sheet sheet = w.createSheet("data");
/*
String NIFTY = driver.findElement(By.id("gettd")).getText();
String NIFTYopen=driver.findElement(By.xpath("//*[@id='dataTable']/tbody/tr[2]/td[4]")).getText();
String NIFTYhigh=driver.findElement(By.xpath("//*[@id=\'dataTable\']/tbody/tr[2]/td[5]")).getText();
String NIFTYlow=driver.findElement(By.xpath("//*[@id=\'dataTable\']/tbody/tr[2]/td[6]")).getText();
String NIFTYltp=driver.findElement(By.xpath("//*[@id=\'dataTable\']/tbody/tr[2]/td[7]")).getText();
String NIFTYchange=driver.findElement(By.xpath("//*[@id=\'dataTable\']/tbody/tr[2]/td[8]")).getText();
String NIFTYperchange=driver.findElement(By.xpath("//*[@id=\'dataTable\']/tbody/tr[2]/td[9]")).getText();
Object[][] bookData = {
        {"Symbol", "Open","High","Low","Last Traded Price","Change","%Change","DV"},
        {NIFTY, NIFTYopen, NIFTYhigh,NIFTYlow,NIFTYltp,NIFTYchange,NIFTYperchange},
        {NIFTY, NIFTYopen, NIFTYhigh,NIFTYlow,NIFTYltp,NIFTYchange,NIFTYperchange},
};
int rowCount = 0;

for (Object[] aBook : bookData) {
    Row row = sheet.createRow(++rowCount);
     
    int columnCount = 0;
     
    for (Object field : aBook) {
        Cell cell = row.createCell(++columnCount);
        if (field instanceof String) {
            cell.setCellValue((String) field);
        } else if (field instanceof Integer) {
            cell.setCellValue((Integer) field);
        }
    }
     
}
*/
List<WebElement> header = driver.findElements(By.xpath("//*[@id=\"dataTable\"]/tbody/tr[1]/th"));
int size = header.size();
for(int i=0;i<size;i++) {
String text = header.get(0).getText();
System.out.println(text);
//List<WebElement> header = driver.findElements(By.xpath("//*[@id=\"dataTable\"]/tbody/tr"));

}


FileOutputStream outputStream = new FileOutputStream("./Excel/NSEdata.xlsx");
w.write(outputStream);
}
}