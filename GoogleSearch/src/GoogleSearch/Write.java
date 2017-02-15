package GoogleSearch;
/*
 * 	This Class will write excel documents and will allow 
 * me to output data in a usable format.
 * 
 * */


import java.io.File;
import java.io.IOException;
import java.util.Locale;
import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


public class Write {
	private WritableCellFormat timesBoldUnderline;
    private WritableCellFormat times;
    private String inputFile;


    public void setOutputFile(String inputFile) {
    
    	this.inputFile = inputFile;
   
    }

    public void write(String[] resultstats, String[] sTerm) throws IOException, WriteException {
    
    	File file = new File(inputFile);
        WorkbookSettings wbSettings = new WorkbookSettings();

        wbSettings.setLocale(new Locale("en", "EN"));

        WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
        workbook.createSheet("Report", 0);
        WritableSheet excelSheet = workbook.getSheet(0);
        createLabel(excelSheet);

        createContent(excelSheet, resultstats, sTerm);

        workbook.write();
        workbook.close();
    }

    private void createLabel(WritableSheet sheet) throws WriteException {

    	// Lets create a times font
        WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);
        // Define the cell format
        times = new WritableCellFormat(times10pt);
        // Lets automatically wrap the cells
        times.setWrap(true);

        // create create a bold font with underlines
        WritableFont times12ptBoldUnderline = new WritableFont( WritableFont.TIMES, 
        		12, WritableFont.BOLD, false, UnderlineStyle.SINGLE);
        
        timesBoldUnderline = new WritableCellFormat(times12ptBoldUnderline);
        
        // Lets automatically wrap the cells
        timesBoldUnderline.setWrap(false);

        CellView cv = new CellView();
        cv.setFormat(times);
        cv.setFormat(timesBoldUnderline);

       // Write a few headers
        addCaption(sheet, 0, 0, "Search term");
        addCaption(sheet, 1, 0, "Results");
        addCaption(sheet, 2, 0, "Time for search");
        
    }

    private void createContent(WritableSheet sheet, String[] resultStats, String[] sTerm) throws WriteException, RowsExceededException {
            
    	String resultTime, resultNum;
    	
    	
    	for (int i = 1; i < resultStats.length; i++) {
    		
    		resultNum = resultStats[i-1].substring(0,(resultStats[i-1].length()-16));
           	resultTime = resultStats[i-1].substring((resultStats[i-1].length()-14),(resultStats[i-1].length()- 10));
           	
        	addLabel(sheet, 0, i, sTerm[i-1]);
        	// First column
            addLabel(sheet, 1, i, resultNum);
            // Second column
            addLabel(sheet, 2, i, resultTime +" seconds");
            // Third column
        }
    
    }

    private void addCaption(WritableSheet sheet, int column, int row, String s) throws RowsExceededException, WriteException {
            
    	Label label;
        label = new Label(column, row, s, timesBoldUnderline);
        sheet.addCell(label);
    }


    private void addLabel(WritableSheet sheet, int column, int row, String s) throws WriteException, RowsExceededException {
        Label label;
        label = new Label(column, row, s, times);
        sheet.addCell(label);
    }

}

