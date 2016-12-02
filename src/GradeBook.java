import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedList;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class GradeBook {
	static final int CAPACITY = 100, ROW_OFFSET = 6, COL_OFFSET = 2, GEN_OFFSET = 4;
	
	public void generate(String fPath) throws IOException {
    	updateOverallGradeSheet();
    	wb.close();
    	
    	try {
			wb.close();
			
	    	FileOutputStream out;
			try {
				out = new FileOutputStream(fPath);
				wb.write(out);
				out.close();
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	// THIS METHOD NEEDS TO INTEGRATE THE GUI EVENT LISTENER
    public void generateCategorySheets(String name, double weight, int numAssign) {
		// Create Category Object and Physical Excel Sheet
		grades.addCategory(name, weight, numAssign);
		Category category = grades.getCategories().getLast();
    	Sheet sheet = wb.createSheet(name);
    	

    	// Populate current category sheet
    	generateCategory(category, sheet);
    }
	
    
    void updateOverallGradeSheet() {
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
    
    
    // THIS METHOD IS FINISHED
    void generateCategory(Category category, Sheet sheet) {
    	generatePointTable(category, sheet);
    	
    	generateInputTable(category, sheet);
    	
    	generateOutputField(category, sheet);
    }
    
    // creates the large input table
    void generateInputTable(Category category, Sheet sheet) {
    	generateStudentTable(sheet);
    	
    	generateInputField(category, sheet);
    	
    	generateOutputField(category, sheet);
    }
    
    void generateOutputField(Category category, Sheet sheet) {
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
    
    void setTotalPointFormula(Row row, Cell cell, int end) {
    	Cell first = row.getCell(COL_OFFSET);
		Cell last = row.getCell(end-1);
		String cellRange = first.getAddress().formatAsString()+":";
		cellRange += last.getAddress().formatAsString();
		cell.setCellFormula("SUM("+cellRange+")");
    }
    
    void setTotalEffectPtFormula(Row row, Cell[] refAssignment, Cell cell) {
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
    
    void setTotalEffectPerFormula(Row row, Cell refTotalPoints, Cell refTotaPercent, Cell cell) {
		StringBuilder formula = new StringBuilder("(");
		Cell effectPts = row.getCell(cell.getColumnIndex()-1);
		formula.append(effectPts.getAddress().formatAsString());
		formula.append("/");
		formula.append(refTotalPoints.getAddress().formatAsString());
		formula.append(")*");
		formula.append(refTotaPercent.getAddress().formatAsString());
		
		cell.setCellFormula(formula.toString());
    }
    
    void setTotalPercentFormula(Row row, Cell cell) {
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
    
    
    void generateInputField(Category category, Sheet sheet) {
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
    void generateStudentTable(Sheet sheet) {
    	Sheet overallGrade = wb.getSheetAt(0);
    	Row currRow;
    	Cell currCell;
    	
    	int r, c;
    	for(r = ROW_OFFSET; r < CAPACITY+ROW_OFFSET+1; ++r) {
    		//if(r == ROW_OFFSET) currRow = sheet.getRow(0); // sheet.createRow(0);
    		//else currRow = sheet.createRow(r);
    		currRow = sheet.createRow(r);
    		for(c = 0; c < COL_OFFSET; ++c) {
        		currCell = currRow.createCell(c);
        		
        		String cellAddress = "'"+overallGrade.getSheetName()+"'!" + 
        		overallGrade.getRow(r - ROW_OFFSET).getCell(c).getAddress().formatAsString();
        		currCell.setCellFormula(cellAddress);
    		}
    	}
    	
    }
    
    void generatePointTable(Category category, Sheet sheet) {
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
    
    void populateTitleRow(Category category, Sheet sheet, Row title) {
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
    
    void populateLabelRow(Category category, Sheet sheet, Row label) {
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
    
    void populateValueRow(Category category, Row value) {
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
    
    void populateDescRow(Category category, Row desc) {
    	Cell label = desc.createCell(COL_OFFSET);
    	label.setCellValue("Description: ");
    	Cell content = desc.createCell(COL_OFFSET+1);
    	content.setCellValue("This is the Reference Table responsible for giving an assignment/ activity its maximum worth (pts).");
    }
    
    
    
    
    // THIS METHOD IS FINISHED
    void generateOverallGradeSheet() {
    	Sheet overallGrade = wb.getSheetAt(0);
    	Row currRow;
    	
    	int r, c;
    	for(r = 0; r < CAPACITY+1; ++r) { // [0, 100]]
    		currRow = overallGrade.createRow(r);
    		for(c = 0; c < 4; ++c) 
    			currRow.createCell(c); // Cell.CELL_TYPE_BLANK
    	}
    	
    	// Generate Column Labels
    	currRow = overallGrade.getRow(0);
    	currRow.getCell(0).setCellValue("SID");
    	currRow.getCell(1).setCellValue("Name");
    	currRow.getCell(2).setCellValue("Effective Total (%)");
    	currRow.getCell(3).setCellValue("Overall Grade (%)");
    }
    
    public GradeBook() {
    	wb = new HSSFWorkbook();
    	grades = new Rubric();
    	
    	wb.createSheet("Overall Grades");
    	generateOverallGradeSheet();
	}
    
    Workbook wb;
    Rubric grades;
}