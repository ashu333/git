package ecom.embeeHub.hub;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
 
public class NsdSupport {
 
        public static void main(String[] args) throws IOException, InterruptedException {
            // System Property for Chrome Driver   
            System.setProperty("webdriver.chrome.driver", "/home/ashish/Downloads/chromedriver_linux64/chromedriver");  
              
                 // Instantiate a ChromeDriver class.     
            WebDriver driver=new ChromeDriver();  
              
        // Launch website  
            driver.navigate().to("https://supporthub.embee.co.in/support/login");  
            driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
            // Click on the search text box and send value  
			 driver.findElement(By.id("username")).sendKeys("javatpoint tutorials");
			 driver.findElement(By.id("password")).sendKeys("passwordials");

			 /* driver.findElement(By.xpath("//html")).click();
			 */  // Click on the search button  
          //  driver.findElement(By.className("header-text btn")).click();  
        }
}