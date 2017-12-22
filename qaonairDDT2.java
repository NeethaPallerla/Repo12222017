import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

import java.io.File;
public class qaonairDDT2 {
	
   WebDriver driver;
   String URL;
   String VFName, VLName, VEmail, VPaswd, VUSERID, Vstatus, VCPaswd, Result;
   String exptxt, acttxt;
   String xlPath, xlSheet,   xlPath_Res,fPath;
   String[][] myTD;
   static int xRows, xCols;
	static int rowCount;
      	   @Before
	
	   public void setup() throws Exception {
		   System.setProperty("webdriver.gecko.driver","C:\\Users\\nspal\\OneDrive\\Documents\\Tools\\gecko\\geckodriver.exe");
		   
		   driver = new FirefoxDriver();
		  // "http://qaonair.com/";
		   
		   xlPath = "C:\\Users\\nspal\\OneDrive\\Documents\\Tools\\DDF1.xls";
		   xlPath_Res = "C:\\Users\\nspal\\OneDrive\\Documents\\Tools\\DDF2_Result.xls";
		   xlSheet = "Sheet1";
		   myTD = readXL(xlPath, xlSheet);
	   }
		   @Test
		   public void signup() throws InterruptedException {
			   System.out.println("Reading the Excel");
			   for (int vrow=1; vrow<xRows; vrow ++) {
				   if (myTD[vrow][1].equals("Y")) {
					   System.out.println("Executing TDID:" + myTD[vrow][0]);
					   VFName = myTD[vrow][2] ;
						 VLName = myTD[vrow][3];
						 VEmail  = myTD[vrow][4]; 
						 VUSERID = myTD[vrow][5];
								 VPaswd= myTD[vrow][6];
								 VCPaswd = myTD[vrow][7];
								 Vstatus = myTD[vrow][8];
								 Vstatus = "Registered user";
								 Vstatus = "New user";
								  driver.get("http://qaonair.com");
								  driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
								   driver.manage().window().maximize();
								
								 enterAUTData();
								 
						   //If user already registered
						   if(Vstatus.equals("Registered user")){
							   Reguser();
						   }
						   //If New user();
						   if(Vstatus.equals("New user")){
							   Newuser();
						   }		   
						   
						   myTD[vrow][10] = acttxt;
						   myTD[vrow][11] = Result;	   
		   
						   
				   }else {
					   System.out.println("Skipping TDID:" + myTD[vrow][0]);
						   
				   }}}
		
		
      //UserDefined Functions
public void enterAUTData() {
	 driver.findElement(By.linkText("SIGN UP")).click();
	   driver.findElement(By.linkText("Sign Up")).click();
	   driver.findElement(By.id("first_name")).sendKeys(VFName);
	   driver.findElement(By.id("last_name")).sendKeys(VLName);
	   driver.findElement(By.id("user_email")).sendKeys(VEmail);
	   driver.findElement(By.id("user_login")).sendKeys(VUSERID);
	   driver.findElement(By.id("user_pass")).sendKeys(VPaswd);
	   driver.findElement(By.id("repeat_pass")).sendKeys(VPaswd);
	  driver.findElement(By.id("signup_form")).click();
	  
	  driver.findElement(By.cssSelector("button.fre-btn.btn-submit")).click();
}
	  
public void Reguser() {
			  exptxt = "Sorry, that username already exists!";
			  acttxt = driver.findElement(By.xpath("/html/body/div[1]/div/div/div/div/form/ul/li")).getText();
			  System.out.println(acttxt);
			  if (exptxt.equals(acttxt)){
				  System.out.println("Test is passed");
				  Result ="Pass";
				  
			  }else {
					  System.out.println("Test is Failed");
					  Result = "Fail";
			  }
			  		  }
			 	   
	   	   public void Newuser() throws InterruptedException {
	  
	   		exptxt = VFName + " " + VLName;
			  acttxt = driver.findElement(By.xpath(".//div[@class='fre-account-info dropdown-toggle']")).getText();
			  System.out.println(acttxt);
			  if (exptxt.equals(acttxt)){
				  System.out.println("Test is passed");
				  Result ="Pass";
				  Thread.sleep(5000);
				  driver.findElement(By.linkText("LOGOUT")).click();
			  }else {
					  System.out.println("Test is Failed");
					  Result ="Fail";
			  }
	   	   }
	   	   
	   	   //Method to Read Excel
	   	   
	   	public static String[][] readXL(String fPath, String fSheet) throws Exception{
			// Purpose : Read an Excel file into a 2D array
			// Inputs : XL Path and XL Sheet name
			// Output : 2D Array
			String[][] xData;  
			
			File myxl = new File(fPath);                               
			FileInputStream myStream = new FileInputStream(myxl);                             
			HSSFWorkbook myWB = new HSSFWorkbook(myStream);                                
			HSSFSheet mySheet = myWB.getSheet(fSheet);                                 
			xRows = mySheet.getLastRowNum()+1;                                
			xCols = mySheet.getRow(0).getLastCellNum();   
			System.out.println("Total Rows in Excel are " + xRows);
			System.out.println("Total Cols in Excel are " + xCols);
			xData = new String[xRows][xCols];        
			for (int i = 0; i < xRows; i++) {                           
				HSSFRow row = mySheet.getRow(i);
				for (int j = 0; j < xCols; j++) {                               
					HSSFCell cell = row.getCell(j);
					String value = "-";
					if (cell!=null){
						value = cell.getStringCellValue();
					}
					xData[i][j] = value;      
					System.out.print(value);
					System.out.print("----");
				}        
				System.out.println("");
			}    
			myxl = null; // Memory gets released
			return xData;
		}
		
	/*	@SuppressWarnings("deprecation")
	public static String cellToString(HSSFCell cell) { 
			// This function will convert an object of type excel cell to a string value
			int type = cell.getCellType();                        
			Object result;                        
			switch (type) {                            
			case HSSFCell.CELL_TYPE_NUMERIC: //0                                
				result = cell.getNumericCellValue();                                
				break;                            
			case HSSFCell.CELL_TYPE_STRING: //1                                
				result = cell.getStringCellValue();                                
				break;                            
			case HSSFCell.CELL_TYPE_FORMULA: //2                                
				throw new RuntimeException("We can't evaluate formulas in Java");  
			case HSSFCell.CELL_TYPE_BLANK: //3                                
				result = "%";                                
				break;                            
			case HSSFCell.CELL_TYPE_BOOLEAN: //4     
				result = cell.getBooleanCellValue();       
				break;                            
			case HSSFCell.CELL_TYPE_ERROR: //5       
				throw new RuntimeException ("This cell has an error");    
			default:                  
				throw new RuntimeException("We don't support this cell type: " + type); 
			}                        
			return result.toString();      
		}*/
		
		public static void writeXL(String fPath, String fSheet, String[][] xData) throws Exception{

			File outFile = new File(fPath);
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet osheet = wb.createSheet(fSheet);
			int xR_TS = xData.length;
			int xC_TS = xData[0].length;
			for (int myrow = 0; myrow < xR_TS; myrow++) {
				HSSFRow row = osheet.createRow(myrow);
				for (int mycol = 0; mycol < xC_TS; mycol++) {
					HSSFCell cell = row.createCell(mycol);
					cell.setCellType(HSSFCell.CELL_TYPE_STRING);
					cell.setCellValue(xData[myrow][mycol]);
				}
				FileOutputStream fOut = new FileOutputStream(outFile);
				wb.write(fOut);
				fOut.flush();
				fOut.close();
			}
		}
		   @After
			  public void teardown() throws Exception {
			   driver.close();
			  writeXL(xlPath_Res, "result_sheet", myTD);
		  }

				
	   
	   }

			   
 
		   
		   
		   
	
		    
		   
		   
	   


     


