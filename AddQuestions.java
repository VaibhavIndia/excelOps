package com.fusion.base;

import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class AddQuestions {

	public static void main(String[] args) throws Exception {

		ReadExcel el = new ReadExcel();
		System.setProperty("webdriver.chrome.driver", "C:\\Selenium\\Sep19\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		
		
		//Login
		driver.get("http://192.168.1.53/fcs/admin/#/app/questionBank");
		driver.findElement(By.name("email")).sendKeys("talentdiagnostic@opencapitaladvisors.com");
		driver.findElement(By.name("password")).sendKeys("123456");
		driver.findElement(By.xpath("//button[@type='submit']")).click();
		
		Thread.sleep(10000);
		
		// Redirect to question bank page
		driver.get("http://192.168.1.53/fcs/admin/#/app/questionBank");
		
		// process to add question
		
		for (int i = 40 ; i < 51 ; i++)
		{
		// click on add question
		driver.findElement(By.xpath("//*[@data-target='#addQuestion']")).click();
		
		// Select category
		
		WebElement categoryDDL = driver.findElement(By.name("section_header"));
		Select select = new Select(categoryDDL);
		String a ="A"+i;
		System.out.println(a);
		select.selectByVisibleText(el.getDatafromParticularCell("Sheet1", a));
		//select.selectByValue(el.getDatafromParticularCell("Sheet1", "A"+i));
		
		// add question text 
		WebElement questionText = driver.findElement(By.xpath("//textarea[@ng-model='question.questionName']"));
		String b ="B"+i;
		System.out.println(b);
		questionText.sendKeys(el.getDatafromParticularCell("Sheet1",b));
		
		//Click on score
		WebElement scoreButton = driver.findElement(By.xpath("//div[@ng-model='question.weight']"));
		scoreButton.click();
		
		// Click on + button for 3 times 
		WebElement plus = driver.findElement(By.xpath("//i[@class='fa fa-plus']"));
		plus.click();
		plus.click();
		plus.click();
		driver.findElement(By.xpath("//div[@class='col-lg-11']/div[2]/div/div/div[2]/input")).sendKeys("Never");
		driver.findElement(By.xpath("//div[@class='col-lg-11']/div[2]/div[2]/input")).sendKeys("1");
		
		driver.findElement(By.xpath("//div[@class='col-lg-11']/div[3]/div/div/div[2]/input")).sendKeys("Sometimes");
		driver.findElement(By.xpath("//div[@class='col-lg-11']/div[3]/div[2]/input")).sendKeys("2");
		
		driver.findElement(By.xpath("//div[@class='col-lg-11']/div[4]/div/div/div[2]/input")).sendKeys("Most of the time");
		driver.findElement(By.xpath("//div[@class='col-lg-11']/div[4]/div[2]/input")).sendKeys("3");
		
		driver.findElement(By.xpath("//div[@class='col-lg-11']/div[5]/div/div/div[2]/input")).sendKeys("Always");
		driver.findElement(By.xpath("//div[@class='col-lg-11']/div[5]/div[2]/input")).sendKeys("4");
		
		driver.findElement(By.xpath("//button[@ng-disabled='submitBtn.disabled']")).click();
		Thread.sleep(10000);
		driver.navigate().refresh();
		
		}

	}

}
