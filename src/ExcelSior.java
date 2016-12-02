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
import javafx.scene.layout.HBox;
import javafx.scene.control.TextArea;
import javafx.scene.paint.Color;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.GridPane;
import javafx.geometry.Insets;
import javafx.scene.layout.*;
import javafx.scene.layout.Background;
//import javafx.scene.layout.BackgroundRepeat;

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.stage.Stage;
import javafx.scene.control.Menu;
import javafx.scene.control.MenuBar;
import javafx.scene.control.MenuItem;
import javafx.geometry.HPos;

import java.util.Scanner;
import java.util.Map;
import java.util.HashMap;
import java.util.LinkedList;
import java.io.File;
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
@SuppressWarnings("restriction")
public class ExcelSior extends Application{
	
	static GradeBook gradeBook;
	/**
	 * This excel spreadsheet generator needs to have the following comprehensive features:
	 * 1)	Parse excel spreadsheets into a GradeDocument
	 * 2)	Write a GradeDocument into an excel spreadsheet
	 */
    public static void main(String[] args) {
    	launch(args);
    }
    
	@Override
	public void start(Stage primaryStage) throws Exception {
		primaryStage.getIcons().add(new Image("Excel-Sior_Logo.png"));
    	primaryStage.setTitle("Excel-Sior");
    	gradeBook = new GradeBook();
    	
    	BorderPane root = new BorderPane();
    	
    	
    	//Image image = new Image(ExcelSior.class.getResource("rsc/Excel-Sior.jpg"));
    	//String path = new File("").getAbsolutePath();
    	//System.out.println(path);
    	//path += "\\rsc\\Excel-Sior.jpg";
    	//rsc/Excel-Sior.jpg
    	
    	//Image img = new Image(getClass().getResourceAsStream("file:/Excel-Sior/rsc/Excel-Sior.jpg"));
    	
    	
    	//String image = ExcelSior.class.getResource("Excel-Sior/rsc/Excel-Sior.jpg").toExternalForm();
    	
    	ImageView img = new ImageView("Excel-Sior.png");
    	//img.setPreserveRatio(true);
    	img.fitWidthProperty().bind(primaryStage.widthProperty());
    	img.fitHeightProperty().bind(primaryStage.heightProperty());
    	root.getChildren().add(img);
    	
    	VBox topContainer = new VBox();
    	
    	//Defining the Menu Bar
    	MenuBar menuBar = generateMenuBar();
        menuBar.prefWidthProperty().bind(primaryStage.widthProperty());
    	topContainer.getChildren().add(menuBar);
    	
    	Scene scene = new Scene(root, 1000, 550, Color.WHITE);
    	root.setTop(topContainer);
    	
        
    	//VBox center = generateCenter();
    	GridPane center = generateGridPane();
    	root.setCenter(center);
    	
    	
    	//root.getCenter().setStyle("-fx-background-image: url(\"/Excel-Sior/rsc/Excel-Sior.jpg\");"
    			//+ "-fx-background-size: 500, 500;-fx-background-repeat: no-repeat;");
    	//
    	
//    	root.setStyle("-fx-background-image: url(\"images/Excel-Sior.png\"); "
//    			+ "-fx-background-position: center center; " 
//    			+ "-fx-background-repeat: stretch;"); // stretch
    	
        primaryStage.setScene(scene);
        primaryStage.show();
	}
	
	static Button generateButton(TextField pathField) {
		Button generate = new Button("Generate");
		generate.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                createNewExcelWorkbook(pathField.getText());
            }
        });
		
		return generate;
	}
	
	
	
	
	
	static GridPane generateGridPane() {
		GridPane grid = new GridPane();
    	grid.setPadding(new Insets(20, 20, 20, 20)); // 10
    	grid.setVgap(5);
    	grid.setHgap(5);
    	//Defining the Name text field
    	final TextField nameField = new TextField();
    	nameField.setPromptText("Enter grade category name.");
    	nameField.setPrefColumnCount(20);
    	nameField.getText();
    	GridPane.setConstraints(nameField, 0, 0);
    	grid.getChildren().add(nameField);
    	//Defining the Last Name text field
    	final TextField weightField = new TextField();
    	weightField.setPromptText("Enter grade category weight.");
    	GridPane.setConstraints(weightField, 0, 1);
    	grid.getChildren().add(weightField);
    	//Defining the Comment text field
    	final TextField numActField = new TextField();
    	numActField.setPrefColumnCount(20);
    	numActField.setPromptText("Enter # of assignments/ activities.");
    	GridPane.setConstraints(numActField, 0, 2);
    	grid.getChildren().add(numActField);
    	
    	TextArea console = addConsole(grid);
    	
    	final TextField pathField = new TextField();
    	pathField.setPrefColumnCount(20);
    	pathField.setPromptText("Enter directory to output Excel file.");
    	GridPane.setConstraints(pathField, 0, 8);
    	grid.getChildren().add(pathField);
    	
    	//Defining the Submit button
    	Button submit = generateAddButton(console, nameField, weightField, numActField);
    	GridPane.setConstraints(submit, 1, 0);
    	grid.getChildren().add(submit);
    	//Defining the Clear button
    	Button clear = generateClearButton(console, nameField, weightField, numActField);
    	GridPane.setConstraints(clear, 1, 1);
    	grid.getChildren().add(clear);
    	//Defining the Clear button
    	Button generate = generateButton(pathField);
    	GridPane.setConstraints(generate, 1, 8);
    	grid.getChildren().add(generate);
    	return grid;
	}
	
	
	
	static TextArea addConsole(GridPane grid){
		//Defining the Label text field
    	final Label label = new Label();
    	label.setText("Console: ");
    	GridPane.setConstraints(label, 0, 5);
    	grid.getChildren().add(label);
    	
    	//Defining the Console text area
    	final TextArea console = new TextArea();
		console.setEditable(false);
		console.setFocusTraversable(false);
    	console.setPrefRowCount(10);
    	console.setPrefColumnCount(100);
    	console.setWrapText(true);
    	console.setPrefWidth(150);
        GridPane.setHalignment(console, HPos.CENTER);
    	GridPane.setConstraints(console, 0, 6);
    	GridPane.setColumnSpan(console, 2);
    	grid.getChildren().add(console);
    	
    	return console;
	}
	
	static Button generateClearButton(TextArea console, TextField nameField, 
			TextField weightField, TextField numActField) {
		Button clear = new Button("Clear");
		 
		//Setting an action for the Clear button
		clear.setOnAction(new EventHandler<ActionEvent>() {
		    public void handle(ActionEvent e) {
		    	nameField.clear();
		    	weightField.clear();
		    	numActField.clear();
		    	console.setText(null);
		    }
		});
		return clear;
	}
	
	
	
	// Change later to a "submit" button
	static Button generateAddButton(TextArea console, TextField nameField, 
			TextField weightField, TextField numActField) {
		Button create = new Button();
    	create.setText("Add Category");
    	create.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
            	StringBuilder msg = new StringBuilder();
            	Double weight = 0.0;
            	Integer numAssign = 0;
            	
            	try {
            		weight = Double.parseDouble(weightField.getText());
            		if(weight <= 0.0) throw new NumberFormatException();
            	} catch(NumberFormatException e) {
            		msg.append("Weight Input Error:\n");
            		msg.append(weightField.getText());
            		msg.append(" is not a valid input.\n");
            	}
            	try {
            		numAssign = Integer.parseInt(numActField.getText());
            		if(numAssign < 1) throw new NumberFormatException();
            	} catch(NumberFormatException e) {
            		msg.append("# Assignment Input Error:\n");
            		msg.append(numActField.getText());
            		msg.append(" is not a valid input.\n");
            	}
            	
            	if(msg.length() == 0) {
            		gradeBook.generateCategorySheets(nameField.getText(), weight, numAssign);
            		msg.append("Successfully Added Category:");
            		msg.append("\n\t(Name):\t\t\t");
            		msg.append(nameField.getText());
            		msg.append("\n\t(Weight):\t\t\t");
            		msg.append(weight);
            		msg.append("\n\t(# Assignments):\t");
            		msg.append(numAssign);
            		msg.append("\n");
            	}
            	
            	console.setText(msg.toString());
            	
            	nameField.clear();
            	weightField.clear();
            	numActField.clear();
            }
        });
		return create;
	}
	
	
	static MenuBar generateMenuBar() {
		MenuBar menuBar = new MenuBar();
		final Menu file = generateFileMenu();
		final Menu options = generateOptionsMenu();
        // MenuItem setting = new MenuItem("setting");
        final Menu help = new Menu("Help");
        
        
        menuBar.getMenus().add(file);
        menuBar.getMenus().add(options);
        menuBar.getMenus().add(help);
		return menuBar;
	}
	
	
	static Menu generateOptionsMenu() {
		Menu options = new Menu("Options");
		MenuItem settings = new MenuItem("Settings");
		
		/* CURRENTLY HAS NO FUNCTIONALITY
		 * THIS IS WHERE THE USER CAN SPECIFY THE DEFAULT OUTPUT DIRECTORY
		options.getItems().add(settings);
        settings.setOnAction(new EventHandler<ActionEvent>() {
            @Override public void handle(ActionEvent e) {
                System.exit(0);
            }
        });*/
        options.getItems().add(settings);
        
        return options;
	}
	
	
	static Menu generateFileMenu() {
		Menu file = new Menu("File");
		file.getItems().add(new MenuItem("Save"));
		file.getItems().add(new MenuItem("Open"));
		file.getItems().add(new SeparatorMenuItem());
        
        MenuItem exit = new MenuItem("Exit");
        exit.setOnAction(new EventHandler<ActionEvent>() {
            @Override public void handle(ActionEvent e) {
                System.exit(0);
            }
        });
        file.getItems().add(exit);
        
        return file;
	}
    
    
    
    static void createNewExcelWorkbook(String fPath) {
    	if(fPath.endsWith("/") || fPath.endsWith("\\")) 
    		fPath += "Fall_2016.xls";
    	else 
    		fPath += "/Fall_2016.xls";
    	
    	try {
			gradeBook.generate(fPath);
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