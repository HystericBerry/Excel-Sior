import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.AbstractMap.SimpleEntry;
import java.util.Map.Entry;
import javafx.scene.control.*;
import javafx.scene.effect.DropShadow;
import javafx.scene.effect.Effect;
import javafx.scene.effect.Glow;
import javafx.scene.effect.SepiaTone;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;


import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.StackPane;
import javafx.stage.Stage;


import java.util.Scanner;
import java.util.Map;
import java.util.List;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.awt.Menu;
import java.awt.MenuBar;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Simple Loan Calculator. Demonstrates advance usage of cell formulas and named ranges.
 *
 * Usage:
 *   LoanCalculator -xls|xlsx
 *
 * @author Yegor Kozlov
 */
public class ExcelSior extends Application{
	private static final int CAPACITY = 100, ROW_OFFSET = 6, COL_OFFSET = 2, GEN_OFFSET = 4;
	
	/**
	 * This excel spreadsheet generator needs to have the following comprehensive features:
	 * 1)	Parse excel spreadsheets into a GradeDocument
	 * 2)	Write a GradeDocument into an excel spreadsheet
	 */
    @SuppressWarnings("restriction")
	public static void main(String[] args) {
    	launch(args);
    }
    
    
	@Override
	@SuppressWarnings("restriction")
	public void start(Stage primaryStage) throws Exception {
		// TODO Auto-generated method stub
    	primaryStage.setTitle("Hello World!");
        Button btn = new Button();
        btn.setText("Create new Excel grade sheet");
        btn.setOnAction(new EventHandler<ActionEvent>() {
 
            @Override
            public void handle(ActionEvent event) {
            	System.out.println(event.toString());
                System.out.println("Hello World!");
            }
        });
        
        
        StackPane root = new StackPane();
        root.getChildren().add(btn);
        primaryStage.setScene(new Scene(root, 300, 250));
        primaryStage.show();
	}
    
    
    static void updateOverallGradeSheet(Workbook wb, Rubric grades) {
    	LinkedList<Category> categories = grades.getCategories();
    	Sheet overallGrade = wb.getSheetAt(0);
		
    	//for every student (row)
    	int i, j;
    	for(i = 1; i < CAPACITY+1; ++i) {
    		Row gradeRow = overallGrade.getRow(i);
    		Cell effPercent = gradeRow.getCell(COL_OFFSET);
    		Cell totalGrade = gradeRow.getCell(COL_OFFSET+1);
    		
    		// for every effPercent & totalGrade column in the other categories
    		StringBuilder perFormula = new StringBuilder();
    		StringBuilder gradeFormula = new StringBuilder("(");
    		for(j = 0; j < categories.size(); ++j) {
    			Category category = categories.get(j);
    			if(j != 0) {
    				perFormula.append(" + ");
    				gradeFormula.append(" + ");
    			}
    			
    			Sheet sheet = wb.getSheetAt(j+1);
    			Row row = sheet.getRow(i+ROW_OFFSET);
    			
    			Cell currPercent = row.getCell(COL_OFFSET+category.getNumAssign()+2); // 2nd to last col
    			Cell currGrade = row.getCell(COL_OFFSET+category.getNumAssign()+3); // last col
        		
    			// e.g. Homework!L4
    			perFormula.append(category.getName());
    			perFormula.append('!');
    			perFormula.append(currPercent.getAddress().formatAsString());
    			
    			// e.g. IF(ISNUMBER(Homework!J4), Homework!J4, 0)
    			String cellAddr = category.getName()+"!"+
    			currGrade.getAddress().formatAsString();
    			
    			gradeFormula.append("IF(ISNUMBER(");
    			gradeFormula.append(cellAddr);
    			gradeFormula.append("),");
    			gradeFormula.append(cellAddr);
    			gradeFormula.append(",0)");
    		}
    		//e.g. Homework!L4 + Project!L4 + Midterm!G4 + ...
    		effPercent.setCellFormula(perFormula.toString());
    		
    		//e.g. (IF(ISNUMBER(Homework!J4), Homework!J4, 0) + ...
    		gradeFormula.append(")/");
    		gradeFormula.append(effPercent.getAddress().formatAsString());
    		totalGrade.setCellFormula(gradeFormula.toString());
    	}
    }
    
    
    // THIS METHOD NEEDS TO INTEGRATE THE GUI EVENT LISTENER
    static void generateCategorySheets(Scanner sc, Workbook wb, Rubric grades) {
    	System.out.println("How many categories?");
    	Integer numCat = readInt(sc);
    	
    	int i;
    	for(i = 0; i < numCat; ++i) {
    		// Implement user input GUI
    		System.out.println("Please enter category name:");
    		String uInput = sc.nextLine();
    		System.out.println("Please enter category weight:");
    		Double weight = readDouble(sc);
    		System.out.println("Please enter # of assignments:");
    		Integer numAssignments = readInt(sc);
    		
    		
    		// Create Category Object and Physical Excel Sheet
    		grades.addCategory(weight, uInput, numAssignments);
    		Category category = grades.getCategories().getLast();
        	Sheet sheet = wb.createSheet(uInput);
        	

        	// Populate current category sheet
        	generateCategory(wb, category, sheet);
    	}
    }
    
    
    // THIS METHOD IS FINISHED
    static void generateCategory(Workbook wb, Category category, Sheet sheet) {
    	generatePointTable(category, sheet);
    	
    	generateInputTable(wb, category, sheet);
    	
    	generateOutputField(category, sheet);
    }
    
    
    /*
    static void processCategories(Scanner sc, Workbook wb, Rubric grades, Integer numCat) {
    	int i;
    	for(i = 1; i <= numCat; ++i) {
    		// Implement user input GUI
    		System.out.println("Please enter category name:");
    		String uInput = sc.nextLine();
    		
    		// Implement user input GUI
    		System.out.println("Please enter category weight:");
    		Double weight = readDouble(sc);
    		
    		System.out.println("Please enter # of assignments:");
    		Integer numAssignments = readInt(sc);
    		
    		grades.addCategory(weight, uInput, numAssignments);
        	wb.createSheet(uInput);
        	
        	generateCategorySheet(sc, grades.getCategory(i-1), wb.getSheetAt(i));
        	setStudentInfo(wb, wb.getSheetAt(i));
        	
        	
        	//Sanity check
        	System.out.println("Category Name: " + wb.getSheetName(i));
        	System.out.println("Category Weight: " + weight);
        	System.out.println("# of sheets: " + wb.getNumberOfSheets());
    	}
    }
    */
    
    // creates the large input table
    static void generateInputTable(Workbook wb, Category category, Sheet sheet) {
    	generateStudentTable(wb, sheet);
    	
    	generateInputField(category, sheet);
    	
    	generateOutputField(category, sheet);
    }
    
    static void generateOutputField(Category category, Sheet sheet) {
    	int i, r, c, start = COL_OFFSET+category.getNumAssign();
    	int end = start + 4;
    	
    	int offset = COL_OFFSET+category.getNumAssign();
    	Row title = sheet.getRow(ROW_OFFSET);
    	Cell totalPoints = title.createCell(offset);
    	sheet.setColumnWidth(offset, 10*256);
    	totalPoints.setCellValue("Total (Pts)");
    	
    	Cell EffectPoints = title.createCell(offset+1);
    	sheet.setColumnWidth(offset+1, 15*256);
    	EffectPoints.setCellValue("Effective (Pts)");
    	
    	Cell EffectPercent = title.createCell(offset+2);
    	sheet.setColumnWidth(offset+2, 15*256);
    	EffectPercent.setCellValue("Effective (%)");
    	
    	Cell totalPercent = title.createCell(offset+3);
    	sheet.setColumnWidth(offset+3, 10*256);
    	totalPercent.setCellValue("Total (%)");
    	
    	
    	// generate the Output Field matrix...
    	final Cell refTotalPoints = sheet.getRow(2).getCell(start);
    	final Cell refTotaPercent = sheet.getRow(2).getCell(start+1);
    	Cell[] refAssignment = new Cell[category.getNumAssign()];
    	for(i = 0; i < refAssignment.length; ++i) 
    		refAssignment[i] = sheet.getRow(2).getCell(i+COL_OFFSET);
    	
    	for(r = ROW_OFFSET+1; r < CAPACITY+ROW_OFFSET+1; ++r) {
    		Row row = sheet.getRow(r);
    		
    		for(c = start; c < end; ++c) {
    			Cell cell = row.createCell(c);
    			if(c == start) 
    				setTotalPointFormula(row, cell, start);
    			else if(c == (start+1)) 
    				setTotalEffectPtFormula(row, refAssignment, cell);
    			else if(c == (start+2)) 
    				setTotalEffectPerFormula(row, refTotalPoints, refTotaPercent, cell);
    			else 
    				setTotalPercentFormula(row, cell);
    		}
    	}
    }
    
    static void setTotalPointFormula(Row row, Cell cell, int end) {
    	Cell first = row.getCell(COL_OFFSET);
		Cell last = row.getCell(end-1);
		String cellRange = first.getAddress().formatAsString()+":";
		cellRange += last.getAddress().formatAsString();
		cell.setCellFormula("SUM("+cellRange+")");
    }
    
    static void setTotalEffectPtFormula(Row row, Cell[] refAssignment, Cell cell) {
    	StringBuilder formula = new StringBuilder();
    	int i;
    	for(i = 0; i < refAssignment.length; ++i) {
    		if(i != 0) formula.append(" + ");
    		
    		
    		StringBuilder ternary = new StringBuilder("IF(ISNUMBER(");
    		ternary.append(row.getCell(i+COL_OFFSET).getAddress().formatAsString());
    		ternary.append("),");
    		ternary.append(refAssignment[i].getAddress().formatAsString());
    		ternary.append(",0)");
    		
    		
    		formula.append(ternary.toString());
    	}
		cell.setCellFormula(formula.toString());
    }
    
    static void setTotalEffectPerFormula(Row row, Cell refTotalPoints, Cell refTotaPercent, Cell cell) {
		StringBuilder formula = new StringBuilder("(");
		Cell effectPts = row.getCell(cell.getColumnIndex()-1);
		formula.append(effectPts.getAddress().formatAsString());
		formula.append("/");
		formula.append(refTotalPoints.getAddress().formatAsString());
		formula.append(")*");
		formula.append(refTotaPercent.getAddress().formatAsString());
		
		cell.setCellFormula(formula.toString());
    }
    
    static void setTotalPercentFormula(Row row, Cell cell) {
    	StringBuilder formula = new StringBuilder("(");
    	Cell totalPts = row.getCell(cell.getColumnIndex()-3);
		Cell effectPts = row.getCell(cell.getColumnIndex()-2);
		Cell effectPer = row.getCell(cell.getColumnIndex()-1);
		
		formula.append(totalPts.getAddress().formatAsString());
		formula.append("/");
		formula.append(effectPts.getAddress().formatAsString());
		formula.append(")*");
		formula.append(effectPer.getAddress().formatAsString());
		
		cell.setCellFormula(formula.toString());
    }
    
    
    
    
    static void generateInputField(Category category, Sheet sheet) {
    	int r, c;
    	Row ref = sheet.getRow(1);
    	Row title = sheet.getRow(ROW_OFFSET);
    	for(c = COL_OFFSET; c < COL_OFFSET+category.getNumAssign(); ++c) {
    		Cell cell = title.createCell(c);
    		cell.setCellFormula(ref.getCell(c).getAddress().formatAsString());
    	}
    	
    	
    	// generate the Input Field matrix...
    	int size = COL_OFFSET+category.getNumAssign(); // +GEN_OFFSET
    	for(r = ROW_OFFSET+1; r < CAPACITY+ROW_OFFSET+1; ++r) {
    		Row row = sheet.getRow(r);
    		for(c = COL_OFFSET; c < size; ++c) 
    			row.createCell(c);
    	}
    }
    
    // THIS METHOD IS FINISHED
    static void generateStudentTable(Workbook wb, Sheet sheet) {
    	Sheet overallGrade = wb.getSheetAt(0);
    	Row currRow;
    	Cell currCell;
    	
    	int r, c;
    	for(r = ROW_OFFSET; r < CAPACITY+ROW_OFFSET+1; ++r) {
    		//if(r == ROW_OFFSET) currRow = sheet.getRow(0); // sheet.createRow(0);
    		//else currRow = sheet.createRow(r);
    		currRow = sheet.createRow(r);
    		for(c = 0; c < COL_OFFSET; ++c) {
        		currCell = currRow.createCell(c, Cell.CELL_TYPE_BLANK);
        		
        		String cellAddress = "'"+overallGrade.getSheetName()+"'!" + 
        		overallGrade.getRow(r - ROW_OFFSET).getCell(c).getAddress().formatAsString();
        		currCell.setCellFormula(cellAddress);
    		}
    	}
    	
    }
    
    
    
    static void generatePointTable(Category category, Sheet sheet) {
    	Row title = sheet.createRow(0);
    	populateTitleRow(category, sheet, title);
    	
    	Row label = sheet.createRow(1);
    	populateLabelRow(category, sheet, label);
    	
    	Row value = sheet.createRow(2);
    	populateValueRow(category, value);
    	
    	Row desc = sheet.createRow(3);
    	populateDescRow(category, desc);
    	
    	sheet.setColumnWidth(COL_OFFSET+category.getNumAssign(), 10*256);
    	sheet.setColumnWidth(COL_OFFSET+category.getNumAssign()+1, 17*256);
    }
    
    static void populateTitleRow(Category category, Sheet sheet, Row title) {
    	// Create N number of columns for Reference Table
    	int i;
    	for(i = COL_OFFSET; i < category.getNumAssign()+GEN_OFFSET; ++i) 
    		title.createCell(i);
    	
    	// Assign Table Name
    	title.getCell(COL_OFFSET).setCellValue("Reference Table (Points)");
    	
    	// Get String Value of Cell Range to Merge
    	StringBuilder cellRange = new StringBuilder();
    	cellRange.append(title.getCell(COL_OFFSET).getAddress().formatAsString());
    	cellRange.append(':');
    	cellRange.append(title.getCell(COL_OFFSET+category.getNumAssign()+1).getAddress().formatAsString());
    	
    	// Merge Title Cells
    	sheet.addMergedRegion(CellRangeAddress.valueOf(cellRange.toString()));
    }
    
    static void populateLabelRow(Category category, Sheet sheet, Row label) {
    	int i;
    	for(i = COL_OFFSET; i < category.getNumAssign()+GEN_OFFSET; ++i) {
    		sheet.setColumnWidth(i, 12*256);
    		Cell cell = label.createCell(i);
    		cell.setCellValue(category.getName()+(i-COL_OFFSET));
    	}
    	
    	Cell totalPts = label.getCell(COL_OFFSET+category.getNumAssign());
    	totalPts.setCellValue("Total (Pts)");
    	Cell gradeWeight = label.getCell(COL_OFFSET+category.getNumAssign()+1);
    	gradeWeight.setCellValue("Grade Weight (Pts)");
    }
    
    static void populateValueRow(Category category, Row value) {
    	int i;
    	for(i = COL_OFFSET; i < category.getNumAssign() + 4; ++i) 
    		value.createCell(i);
    	
    	Cell totalPts = value.getCell(COL_OFFSET+category.getNumAssign());
    	StringBuilder cellRange = new StringBuilder();
    	cellRange.append(value.getCell(COL_OFFSET).getAddress().formatAsString());
    	cellRange.append(":");
    	cellRange.append(value.getCell(COL_OFFSET+category.getNumAssign()-1).getAddress().formatAsString());
    	totalPts.setCellFormula("SUM("+cellRange.toString()+")");
    	
    	Cell gradeWeight = value.getCell(COL_OFFSET+category.getNumAssign()+1);
    	gradeWeight.setCellValue(category.getWeight());
    }
    
    static void populateDescRow(Category category, Row desc) {
    	Cell label = desc.createCell(COL_OFFSET);
    	label.setCellValue("Description: ");
    	Cell content = desc.createCell(COL_OFFSET+1);
    	content.setCellValue("This is the Reference Table responsible for giving an assignment/ activity its maximum worth (pts).");
    }
    
    
    
    
    // THIS METHOD IS FINISHED
    static void generateOverallGradeSheet(Workbook wb) {
    	Sheet overallGrade = wb.getSheetAt(0);
    	Row currRow;
    	Cell currCell;
    	
    	int r, c;
    	for(r = 0; r < CAPACITY+1; ++r) { // [0, 100]]
    		currRow = overallGrade.createRow(r);
    		for(c = 0; c < 4; ++c) 
    			currCell = currRow.createCell(c); // Cell.CELL_TYPE_BLANK
    	}
    	
    	// Generate Column Labels
    	currRow = overallGrade.getRow(0);
    	currRow.getCell(0).setCellValue("SID");
    	currRow.getCell(1).setCellValue("Name");
    	currRow.getCell(2).setCellValue("Effective Total (%)");
    	currRow.getCell(3).setCellValue("Overall Grade (%)");
    }
    
    
    static Integer readInt(Scanner sc) {
    	String uInput;
    	Integer numCat;
    	
    	while (true) { // Parse Integer
    		uInput = sc.nextLine();
    		try {
    			numCat = Integer.parseInt(uInput);
    		} catch (NumberFormatException e) {
    			System.out.println("Please try again lulz.");
    			continue;
    		}
    		break;
    	}
    	
    	return numCat;
    }
    
    static Double readDouble(Scanner sc) {
    	String uInput;
    	Double numCat;
    	
    	while (true) { // Parse Integer
    		uInput = sc.nextLine();
    		try {
    			numCat = Double.parseDouble(uInput);
    		} catch (NumberFormatException e) {
    			System.out.println("Please try again lulz.");
    			continue;
    		}
    		break;
    	}
    	
    	return numCat;
    }
    
    
    static void createNewExcelWorkbook() {
    	String fPath = "C:/Users/pkim7/Desktop/output/Fall_2016.xls";
    	Scanner sc = new Scanner(System.in);

    	Workbook wb = new HSSFWorkbook();
    	Rubric grades = new Rubric();
    	
    	System.out.println("Initializing workbook:");
    	wb.createSheet("Overall Grades");
    	generateOverallGradeSheet(wb);
    	System.out.println("# of sheets: " + wb.getNumberOfSheets() + "\n");
    	
    	generateCategorySheets(sc, wb, grades);
    	updateOverallGradeSheet(wb, grades);
    	
    	try {
			wb.close();
			
			System.out.println("Creating Excel File...");
	    	FileOutputStream out;
			try {
				out = new FileOutputStream(fPath);
				wb.write(out);
				out.close();
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
	    	System.out.println("Program exited");
	    	sc.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
    
    /**
     * (Future Implementation)
     * Given a String directory to the excel file, this method reads and populates categories.
     */
    static void parseExcelSpreadsheet() {
    	
    }

    //define named ranges for the inputs and formulas
    static void createNames(Workbook wb){
        Name name;

        name = wb.createName();
        name.setNameName("Interest_Rate");
        name.setRefersToFormula("'Loan Calculator'!$E$5");

        name = wb.createName();
        name.setNameName("Loan_Amount");
        name.setRefersToFormula("'Loan Calculator'!$E$4");

        name = wb.createName();
        name.setNameName("Loan_Start");
        name.setRefersToFormula("'Loan Calculator'!$E$7");

        name = wb.createName();
        name.setNameName("Loan_Years");
        name.setRefersToFormula("'Loan Calculator'!$E$6");

        name = wb.createName();
        name.setNameName("Number_of_Payments");
        name.setRefersToFormula("'Loan Calculator'!$E$10");

        name = wb.createName();
        name.setNameName("Monthly_Payment");
        name.setRefersToFormula("-PMT(Interest_Rate/12,Number_of_Payments,Loan_Amount)");

        name = wb.createName();
        name.setNameName("Total_Cost");
        name.setRefersToFormula("'Loan Calculator'!$E$12");

        name = wb.createName();
        name.setNameName("Total_Interest");
        name.setRefersToFormula("'Loan Calculator'!$E$11");

        name = wb.createName();
        name.setNameName("Values_Entered");
        name.setRefersToFormula("IF(Loan_Amount*Interest_Rate*Loan_Years*Loan_Start>0,1,0)");
    }
    
    
    /**
     * cell styles used for formatting calendar sheets
     */
    static Map<String, CellStyle> createStyles(Workbook wb){
        Map<String, CellStyle> styles = new HashMap<String, CellStyle>();

        CellStyle style;
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short)14);
        titleFont.setFontName("Trebuchet MS");
        style = wb.createCellStyle();
        style.setFont(titleFont);
        style.setBorderBottom(BorderStyle.DOTTED);
        style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        styles.put("title", style);

        Font itemFont = wb.createFont();
        itemFont.setFontHeightInPoints((short)9);
        itemFont.setFontName("Trebuchet MS");
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setFont(itemFont);
        styles.put("item_left", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setFont(itemFont);
        styles.put("item_right", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setFont(itemFont);
        style.setBorderRight(BorderStyle.DOTTED);
        style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.DOTTED);
        style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.DOTTED);
        style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.DOTTED);
        style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setDataFormat(wb.createDataFormat().getFormat("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"));
        styles.put("input_$", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setFont(itemFont);
        style.setBorderRight(BorderStyle.DOTTED);
        style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.DOTTED);
        style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.DOTTED);
        style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.DOTTED);
        style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setDataFormat(wb.createDataFormat().getFormat("0.000%"));
        styles.put("input_%", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setFont(itemFont);
        style.setBorderRight(BorderStyle.DOTTED);
        style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.DOTTED);
        style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.DOTTED);
        style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.DOTTED);
        style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setDataFormat(wb.createDataFormat().getFormat("0"));
        styles.put("input_i", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFont(itemFont);
        style.setDataFormat(wb.createDataFormat().getFormat("m/d/yy"));
        styles.put("input_d", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setFont(itemFont);
        style.setBorderRight(BorderStyle.DOTTED);
        style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.DOTTED);
        style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.DOTTED);
        style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.DOTTED);
        style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setDataFormat(wb.createDataFormat().getFormat("$##,##0.00"));
        style.setBorderBottom(BorderStyle.DOTTED);
        style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put("formula_$", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setFont(itemFont);
        style.setBorderRight(BorderStyle.DOTTED);
        style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.DOTTED);
        style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.DOTTED);
        style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.DOTTED);
        style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setDataFormat(wb.createDataFormat().getFormat("0"));
        style.setBorderBottom(BorderStyle.DOTTED);
        style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put("formula_i", style);

        return styles;
    }
}




/*
static void init() {
	Workbook wb = new HSSFWorkbook();
    

    //if(args.length > 0 && args[0].equals("-xls")) wb = new HSSFWorkbook();
    //else wb = new XSSFWorkbook();

    
    wb.getSheetAt(0);
    
    Map<String, CellStyle> styles = createStyles(wb); // getNumberOfSheets
    Sheet sheet = wb.createSheet("Overall Grade Sheet");
    sheet.setPrintGridlines(false);
    sheet.setDisplayGridlines(false);

    PrintSetup printSetup = sheet.getPrintSetup();
    printSetup.setLandscape(true); // do i need this?
    sheet.setFitToPage(true);
    sheet.setHorizontallyCenter(true);

    sheet.setColumnWidth(0, 3*256);
    sheet.setColumnWidth(1, 3*256);
    sheet.setColumnWidth(2, 11*256);
    sheet.setColumnWidth(3, 14*256);
    sheet.setColumnWidth(4, 14*256);
    sheet.setColumnWidth(5, 14*256);
    sheet.setColumnWidth(6, 14*256);

    createNames(wb);

    Row titleRow = sheet.createRow(0);
    titleRow.setHeightInPoints(35);
    for (int i = 1; i <= 7; i++) { // column
        titleRow.createCell(i).setCellStyle(styles.get("title"));
    }
    Cell titleCell = titleRow.getCell(2);
    titleCell.setCellValue("Simple Loan Calculator");
    sheet.addMergedRegion(CellRangeAddress.valueOf("$C$1:$H$1"));

    


    // Write the output to a file
    String dir = "C:/Users/pkim7/Desktop/GradeSheet.xls"; // D:/Paul_Kim/Documents/GradeSheet.xls
    if(wb instanceof XSSFWorkbook) dir += "x";
    FileOutputStream out = new FileOutputStream(dir);
    wb.write(out);
    out.close();
}*/


/*
Row row = sheet.createRow(2);
Cell cell = row.createCell(4);
cell.setCellValue("Enter values");
cell.setCellStyle(styles.get("item_right"));

row = sheet.createRow(3);
cell = row.createCell(2);
cell.setCellValue("Loan amount");
cell.setCellStyle(styles.get("item_left"));
cell = row.createCell(4);
cell.setCellStyle(styles.get("input_$"));
cell.setAsActiveCell();

row = sheet.createRow(4);
cell = row.createCell(2);
cell.setCellValue("Annual interest rate");
cell.setCellStyle(styles.get("item_left"));
cell = row.createCell(4);
cell.setCellStyle(styles.get("input_%"));

row = sheet.createRow(5);
cell = row.createCell(2);
cell.setCellValue("Loan period in years");
cell.setCellStyle(styles.get("item_left"));
cell = row.createCell(4);
cell.setCellStyle(styles.get("input_i"));

row = sheet.createRow(6);
cell = row.createCell(2);
cell.setCellValue("Start date of loan");
cell.setCellStyle(styles.get("item_left"));
cell = row.createCell(4);
cell.setCellStyle(styles.get("input_d"));

row = sheet.createRow(8);
cell = row.createCell(2);
cell.setCellValue("Monthly payment");
cell.setCellStyle(styles.get("item_left"));
cell = row.createCell(4);
cell.setCellFormula("IF(Values_Entered,Monthly_Payment,\"\")");
cell.setCellStyle(styles.get("formula_$"));

row = sheet.createRow(9);
cell = row.createCell(2);
cell.setCellValue("Number of payments");
cell.setCellStyle(styles.get("item_left"));
cell = row.createCell(4);
cell.setCellFormula("IF(Values_Entered,Loan_Years*12,\"\")");
cell.setCellStyle(styles.get("formula_i"));

row = sheet.createRow(10);
cell = row.createCell(2);
cell.setCellValue("Total interest");
cell.setCellStyle(styles.get("item_left"));
cell = row.createCell(4);
cell.setCellFormula("IF(Values_Entered,Total_Cost-Loan_Amount,\"\")");
cell.setCellStyle(styles.get("formula_$"));

row = sheet.createRow(11);
cell = row.createCell(2);
cell.setCellValue("Total cost of loan");
cell.setCellStyle(styles.get("item_left"));
cell = row.createCell(4);
cell.setCellFormula("IF(Values_Entered,Monthly_Payment*Number_of_Payments,\"\")");
cell.setCellStyle(styles.get("formula_$"));
*/