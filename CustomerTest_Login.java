import java.io.File;
import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class CustomerTest_Login {

	public static void main(String[] args) throws BiffException, IOException, InterruptedException {
		// TODO Auto-generated method stub
		FirefoxDriver driver = new FirefoxDriver();
		driver.get("http://alaldinf131.risk.regn.net/");
		String title = driver.getTitle();
		System.out.println(title);
		driver.findElement(By.xpath("//*[@id='login']/div[1]/input")).sendKeys("sujata.senadhi@lexisnexis.com");
		driver.findElement(By.xpath("//*[@id='login']/div[2]/input")).sendKeys("98rquCDo");
		driver.findElement(By.xpath("//*[@id='login']/a/div")).click();
		driver.findElement(By.xpath("//*[@id='vertmenu']/ul/li[1]/ul/li[1]/a")).click();

		String products[] = { "ADD" };

		Workbook excel = Workbook.getWorkbook(new File("PATHOFEXCELFILE"));
		Sheet locations = excel.getSheet("location");

		Select prouctDropDown = new Select(driver.findElement(By.xpath(".//*[@id='product']")));
		Select state = new Select(driver.findElement(By.xpath(".//*[@name='state']")));
		for (int i = 0; i < products.length; i++) {
			//Select product
			prouctDropDown.selectByVisibleText(products[i]);

			for (int j = 0; j < locations.getRows(); j++) {
			//select location
				state.selectByVisibleText(locations.getCell(0, j).getContents());
			}

			Thread.sleep(4000);
			driver.findElement(By.id("downloaddata")).click();
		}

		driver.findElement(By.xpath("//*[@id='td']")).click();

	}

}
