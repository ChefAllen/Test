package GoogleSearch;
/*
 * this reads an excel document and returns the items as a string array
 * this array will be used to do the searches for the test 
 */


import java.io.File;
import java.io.IOException;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class Read {

	private String inputFile;
	
    public void setInputFile(String inputFile) {
    	
    	this.inputFile = inputFile;
        
    }
    
    public String[] read() throws IOException  {
                
    	File inputWorkbook = new File(inputFile);
        Workbook w;
        
        try {            
        	w = Workbook.getWorkbook(inputWorkbook);
        	
        	Sheet sheet = w.getSheet(0);
        	String[] a = new String[sheet.getColumns()*sheet.getRows()];
        	
            for (int j = 0; j < sheet.getColumns(); j++) {
                                
            	for (int i = 1; i < sheet.getRows(); i++) {
                                        
            		Cell cell = sheet.getCell(j, i);
                    a[(i+j)-1] = (String)cell.getContents();

            	}        
            }
            return a;
                
        } catch (BiffException e) {
                        
        	e.printStackTrace();
        	return null;
                
        }
        
    }
	
}
