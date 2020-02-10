package utilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Webdriver;

public class ReadExcel extends PageBase {
  
  public ReadExcel(WebDriver driver) {
    super(driver);
    //TODO Auto-generated constructor stub
  }
  
  public ReadExcel loadAndReadExcel(String fileLocation, String fileName) throws IOException {
    File source = new File(fileLocation + "\\" + fileName);
    FileInputStream fileInputStream = new FileInputStream(source);
    XSSFWorkbook workBook = new XSSFWorkbook(fileInputStream);
    XSSFSheet sheet = workBook.getSheetAt(0);
    
    System.out.printLn(sheet.getRow(15).getCell(3).getStringCellValue());
    
    return this;
  }
}
  
  
  
  
