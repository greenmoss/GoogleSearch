/*
 * this application is made to take an excel sheet of terms and go through it as a list, it then searches google on 3 separate browsers(IE, Chrome, Firefox) and returns
 * information about what it finds, currently it finds the number of results given by google and the seconds taken to find those results. it then save a screencap named
 * for the term it's searching for, finally the application parses the information into a new excel document which is generated in the same folder as the screen shots.
 */


package GoogleSearch;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

import javax.imageio.ImageIO;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;

import jxl.write.WriteException;

public class Webwork {

    public static void FFLaunch(String[] a) throws IOException {
    	
        // Firefox browser setup
        System.setProperty("webdriver.gecko.driver","C:\\Users\\Yggdrasil\\Desktop\\coding\\Selenium\\jar files and resources\\geckodriver.exe");
        DesiredCapabilities capabilities = DesiredCapabilities.firefox();
        capabilities.setCapability("marionette", true);
        
        //---instantiate objects
        WebDriver FFdriver = new FirefoxDriver();
        Write output = new Write();
        String browser = "FF";
        
        //---run test
        String[] results = Test(FFdriver, a, browser);
        
        //--generate excel documents with findings
        try{
			output.setOutputFile("C:\\Users\\Yggdrasil\\Desktop\\coding\\Selenium\\ScreenCaps\\" + browser +"\\results.xls");
			output.write(results,a);
		
		} catch(WriteException e){
			
			System.out.println("there was an issue");
		
		}
            
    }
   
    public static void ChromeLaunch(String[] a) throws IOException{
        //Chrome browser setup
        System.setProperty("webdriver.chrome.driver","C:\\Users\\Yggdrasil\\Desktop\\coding\\Selenium\\jar files and resources\\chromedriver.exe");
        
        //---instantiate objects
        WebDriver ChromeDriver = new ChromeDriver();
        Write output = new Write();
        String browser = "Chrome";
        
        //---run test
        String[] results = Test(ChromeDriver, a, browser);
        
        
        try{
			output.setOutputFile("C:\\Users\\Yggdrasil\\Desktop\\coding\\Selenium\\ScreenCaps\\" + browser +"\\results.xls");
			output.write(results,a);
		
		} catch(WriteException e){
			
			System.out.println("there was an issue");
		
		}
            
    }
        
   
    public static void IELaunch(String[] a) throws IOException{
        //IE browser setup
        File file = new File("C:\\Users\\Yggdrasil\\Desktop\\coding\\Selenium\\jar files and resources\\IEDriverServer32.exe");
        System.setProperty("webdriver.ie.driver", file.getAbsolutePath());
        
        //---instantiate objects
        WebDriver IEDriver = new InternetExplorerDriver();
        Write output = new Write();
        String browser = "IE";
        String[] results = Test(IEDriver, a, browser);
        
        //--generate excel documents with findings
        try{
			output.setOutputFile("C:\\Users\\Yggdrasil\\Desktop\\coding\\Selenium\\ScreenCaps\\" + browser +"\\results.xls");
			output.write(results,a);
		
		} catch(WriteException e){
			
			System.out.println("there was an issue");
		
		}
            
    }
    
    
    
    public static void ScreenCap(String a, String browser){
    	
    	
    	//---Captures the entire screen and saves it in the appropriate file for the browser as a .png with appropriate name
    	
    	Rectangle screenRect = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
    	
    	try{
        		   
    		Robot image = new Robot();
        	BufferedImage stuff = null;
        	stuff = image.createScreenCapture(new Rectangle(screenRect));
        	   
        	File output = new File("C:\\Users\\Yggdrasil\\Desktop\\coding\\Selenium\\ScreenCaps\\" + browser +"\\" + a + ".png");
        	ImageIO.write(stuff, "png", output);
        	
    	} catch (Exception e) {
        		   
    		System.out.println("issue with ScreenCap.");
        	   
    	}
    }
    
    
    public static String[] Test(WebDriver driver, String[] a, String browser){
    	
    	//---instantiate variables and open browser
    	String resultStats, resultTime, resultNum;
    	String[] resultOut = new String[a.length];
    	driver.get("http://www.google.com");
    	
       try{
            Thread.sleep(2500);
        } catch (InterruptedException e){
           
        }
       
       for(int i = 0; i< (a.length - 1); i++){
       
    	   //--- navigate through search engine with data taken from document
    	   driver.findElement(By.id("lst-ib")).clear();
    	   driver.findElement(By.id("lst-ib")).sendKeys(a[i]);
    	   driver.findElement(By.xpath("//button[@ id='_fZl']")).click(); 
       	   
       	try{
            Thread.sleep(2000);
        } catch (InterruptedException e){
           
        }
       	
       	//--- collect and parse out data
       	resultStats = (String) driver.findElement(By.xpath("//div[@ id='resultStats']")).getText();
       	resultOut[i]=resultStats;
       	resultNum = resultStats.substring(0,(resultStats.length()-16));
       	resultTime = resultStats.substring((resultStats.length()-14),(resultStats.length()- 10));
       	
       	//read back data collected to console
    	System.out.println("a[" + i + "] which is " + a[i] + " has " + resultNum + ", it took about " + resultTime +" seconds.");
    	
    	//capture screen and save picture
    	ScreenCap(a[i],browser);
    	   
       }
       
   	   //---close browser
       driver.quit();
    	
       //--- return findings
       return resultOut;
    }
    
    
    public static void main(String args[]) throws Exception{
		
		//Instantiate Objects
		Read test = new Read();
		
		//---bringing the items from the excel document for searching
        test.setInputFile("C:\\Users\\Yggdrasil\\Desktop\\coding\\Selenium\\ExcelFiles\\SearchTerms.xls");
        String[] testData = test.read();
        
        //---verify testdata present
        for(int i = 0; i < (testData.length - 1); i++){
        	System.out.println("testData[" + i + "] = " + testData[i]);
        }
        
        //--- Call each browser
		Webwork.FFLaunch(testData);
		Webwork.ChromeLaunch(testData);
		Webwork.IELaunch(testData);
		
		System.out.println("Complete!");
	}

}