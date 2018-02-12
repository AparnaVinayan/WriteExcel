package org.all.WriteExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class ExcelWrite {
	public static void main(String[] args) throws IOException {
		 System.setProperty("webdriver.chrome.driver",
					"C:/Users/Aparna/eclipse-workspace/TrialFB/Drivers/chromedriver.exe");
		 WebDriver driver = new ChromeDriver();
		driver.get("https://www.nseindia.com/live_market/dynaContent/live_watch/equities_stock_watch.htm?cat=N"); 
		Workbook w=new XSSFWorkbook();
		Sheet sheet = w.createSheet("Java Books");
        
        Object[][] bookData = {
                {"Head First Java", "Kathy Serria", 79},
                {"Effective Java", "Joshua Bloch", 36},
                {"Clean Code", "Robert martin", 42},
                {"Thinking in Java", "Bruce Eckel", 35},
                {"acv", "avbn", 13}
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
       FileOutputStream outputStream = new FileOutputStream("./Excel/JavaBooks.xlsx");
            w.write(outputStream);
        


	  }	

}
