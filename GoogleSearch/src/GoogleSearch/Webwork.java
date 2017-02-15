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

    public static void FFLaunch(String[] testData, String mainFolder) throws IOException {
    	
        // Firefox browser setup
    	File file = new File(System.getProperty("user.home") + "\\Desktop\\Drivers\\geckodriver.exe");
        System.setProperty("webdriver.gecko.driver", file.getAbsolutePath());
        DesiredCapabilities capabilities = DesiredCapabilities.firefox();
        capabilities.setCapability("marionette", true);
        
        //---instantiate objects
        WebDriver FFdriver = new FirefoxDriver();
        Write output = new Write();
        String browser = "Firefox";
        
        //---make directory
        FolderMaker(mainFolder + "\\" + browser);
        
        //---run test
        String[] results = Test(FFdriver, testData, browser);
        
        //--generate excel documents with findings
        try{
        	output.setOutputFile(System.getProperty("user.home") + "\\Desktop\\Results\\" + browser +"\\results.xls");
        	output.write(results, testData);
		
		} catch(WriteException e){
			
			System.out.println("there was an issue");
		
		}
            
    }
   
    public static void ChromeLaunch(String[] testData, String mainFolder) throws IOException{
        //Chrome browser setup
    	File file = new File(System.getProperty("user.home") + "\\Desktop\\Drivers\\chromedriver.exe");
        System.setProperty("webdriver.chrome.driver", file.getAbsolutePath());
        //---instantiate objects
        WebDriver ChromeDriver = new ChromeDriver();
        Write output = new Write();
        String browser = "Chrome";
        
      //---make directory
        FolderMaker(mainFolder + "\\" + browser);
        
        //---run test
        String[] results = Test(ChromeDriver, testData, browser);
        
        
        try{
        	output.setOutputFile(System.getProperty("user.home") + "\\Desktop\\Results\\" + browser +"\\results.xls");
        	output.write(results,testData);
		
		} catch(WriteException e){
			
			System.out.println("there was an issue");
		
		}
            
    }
        
   
    public static void IELaunch(String[] testData, String mainFolder) throws IOException{
        //IE browser setup
        File file = new File(System.getProperty("user.home") + "\\Desktop\\Drivers\\IEDriverServer32.exe");
        System.setProperty("webdriver.ie.driver", file.getAbsolutePath());
        
        //---instantiate objects
        WebDriver IEDriver = new InternetExplorerDriver();
        Write output = new Write();
        String browser = "InternetExplorer";
        
      //---make directory
        FolderMaker(mainFolder + "\\" + browser);
        
        //---run test
        String[] results = Test(IEDriver, testData, browser);
        
        //--generate excel documents with findings
        try{
			output.setOutputFile(System.getProperty("user.home") + "\\Desktop\\Results\\" + browser +"\\results.xls");
			output.write(results,testData);
		
		} catch(WriteException e){
			
			System.out.println("there was an issue");
		
		}
            
    }
    
    public static void FolderMaker(String fileName){
		String desktop = System.getProperty("user.home") + "\\Desktop";
		System.out.println(desktop + "\\" + fileName + " created");
		File dir = new File(desktop + "\\" + fileName);
		dir.mkdirs();
	}
    
    public static void ScreenCap(String a, String browser){
    	
    	
    	//---Captures the entire screen and saves it in the appropriate file for the browser as a .png with appropriate name
    	
    	Rectangle screenRect = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
    	
    	try{
        		   
    		Robot robot = new Robot();
        	BufferedImage image = null;
        	image = robot.createScreenCapture(new Rectangle(screenRect));
        	   
        	File output = new File(System.getProperty("user.home") + "\\Desktop\\Results\\" + browser +"\\" + a + ".png");
        	ImageIO.write(image, "png", output);
        	
    	} catch (Exception e) {
        		   
    		System.out.println("issue with ScreenCap.");
        	   
    	}
    }
    
    
    public static String[] Test(WebDriver driver, String[] testData, String browser){
    	
    	//---instantiate variables and open browser
    	String resultStats, resultTime, resultNum;
    	String[] resultOut = new String[testData.length];
    	driver.get("http://www.google.com");
    	
       try{
            Thread.sleep(2500);
        } catch (InterruptedException e){
           
        }
       
       for(int i = 0; i< (testData.length - 1); i++){
       
    	   //--- navigate through search engine with data taken from document
    	   driver.findElement(By.id("lst-ib")).clear();
    	   driver.findElement(By.id("lst-ib")).sendKeys(testData[i]);
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
    	System.out.println("testData[" + i + "] has " + resultNum + ", it took about " + resultTime +" seconds.");
    	
    	//capture screen and save picture
    	ScreenCap(testData[i],browser);
    	   
       }
       
   	   //---close browser
       driver.quit();
    	
       //--- return findings
       return resultOut;
    }
    
    
    public static void main(String args[]) throws Exception{
		
		//Instantiate Objects
		Read test = new Read();
		String mainFolder = "Results";
		
		//---bringing the items from the excel document for searching
        test.setInputFile(System.getProperty("user.home") + "\\Desktop\\SearchTerms.xls");
        String[] testData = test.read();
        
        //---verify testdata present
        for(int i = 0; i < (testData.length - 1); i++){
        	System.out.println("testData[" + i + "] = " + testData[i]);
        }
        FolderMaker(mainFolder);
        
        //--- Call each browser
		Webwork.FFLaunch(testData, mainFolder);
		Webwork.ChromeLaunch(testData, mainFolder);
		Webwork.IELaunch(testData, mainFolder);
		
		System.out.println("Complete!");
	}

}
