import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class UniApp
{
	static WebDriver driver;
	static List<WebElement> elements;
	static XSSFSheet sheet;
	static int i = 0;
	
	public static void main(String[] args) throws InterruptedException, IOException
	{
		System.setProperty("webdriver.chrome.driver", "./chromedriver.exe");
		
		driver = new ChromeDriver();
		driver.get(
				"https://www.usnews.com/best-graduate-schools/top-science-schools/computer-science-rankings");
		
		loadByScrolling();
		
		// EXCEL-----------
		File src = new File("./test.xlsx");
		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		sheet = wb.getSheetAt(0);
		// ----------------
		
		elements = driver.findElements(By.className("kKdFhD"));
		
		writeDataToExcel();
		
		FileOutputStream fout = new FileOutputStream(src);
		wb.write(fout);
		wb.close();
	}
	
	/**
	 * Writes the names, rank, and location(state) of all universities from the
	 * web page to an excel file.
	 */
	private static void writeDataToExcel()
	{
		for (WebElement el : elements)
		{
			String uniName = el.findElement(By.tagName("a")).getText();
			String uniRank = el.findElement(By.className("eULIZs")).getText();
			String uniState = el.findElement(By.className("fJtpNK")).getText();
			
			if (!uniName.isEmpty())
			{
				System.out.println(": " + uniName);
				sheet.createRow(i);
				sheet.getRow(i).createCell(0).setCellValue(uniName);
				System.out.println(uniRank);
				sheet.getRow(i).createCell(1).setCellValue(uniRank);
				System.out.println(uniState);
				sheet.getRow(i).createCell(2).setCellValue(uniState);
				i++;
			}
		}
	}
	
	/**
	 * Scrolls the page to trigger the dynamically loading information about
	 * universities
	 * 
	 * @throws InterruptedException
	 */
	private static void loadByScrolling() throws InterruptedException
	{
		WebElement loadBtn = driver.findElement(By.className("BRQzn"));
		JavascriptExecutor js = (JavascriptExecutor) driver;
		
		for (int i = 0; i < 10; i++)
		{
			try
			{
				js.executeScript("arguments[0].scrollIntoView();", loadBtn);
				if (i > 3)
				{
					loadBtn.click();
				}
			}
			catch (StaleElementReferenceException e)
			{
				break;
			}
			catch (Exception e)
			{
				i--;
			}
			
			System.out.println("SCROLLED");
			Thread.sleep(3500);
		}
	}
}
