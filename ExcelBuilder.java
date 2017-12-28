package com.beingjavaguys.utility;

import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.springframework.web.servlet.view.document.AbstractExcelView;

import com.beingjavaguys.domain.Student;


public class ExcelBuilder extends AbstractExcelView {


	@Override
	protected void buildExcelDocument(Map<String, Object> model,
			HSSFWorkbook workbook, HttpServletRequest request, HttpServletResponse response)
					throws Exception {
		// get data model which is passed by the Spring container
		Map<String,Object> map = (Map<String,Object>) model.get("map");

		List<Student> studentList = (List<Student>) map.get("student");

		// create a new Student sheet
		HSSFSheet studentSheet = workbook.createSheet("Student");
		studentSheet.setDefaultColumnWidth(15);

		HSSFRow aRow = studentSheet.createRow(0);
		aRow.createCell(0).setCellValue("report type");
		aRow.createCell(1).setCellValue("Pre-generated report");

		HSSFRow aRow1 = studentSheet.createRow(1);
		aRow1.createCell(1).setCellValue("Monthly");

		HSSFRow aRow2 = studentSheet.createRow(2);
		aRow2.createCell(1).setCellValue("Business Area");

		HSSFRow aRow3 = studentSheet.createRow(3);
		aRow3.createCell(1).setCellValue("Department");

		// create style for Row 6
		CellStyle style6 = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setFontName("Arial");
		style6.setFillForegroundColor(HSSFColor.GREY_80_PERCENT.index);
		style6.setFillPattern(CellStyle.SOLID_FOREGROUND);
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setColor(HSSFColor.WHITE.index);
		style6.setFont(font);
		style6.setAlignment(CellStyle.ALIGN_CENTER);

		studentSheet.addMergedRegion(new CellRangeAddress(6, 6, 0, 10));
		HSSFRow aRow6 = studentSheet.createRow(6);

		aRow6.createCell(0).setCellValue("Idea Share Annual Report");
		aRow6.getCell(0).setCellStyle(style6);
		aRow6.setHeight((short) 600);

		// create style for Row 7
		CellStyle style7 = workbook.createCellStyle();
		Font font7 = workbook.createFont();
		font7.setFontName("Arial");
		style7.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
		style7.setFillPattern(CellStyle.SOLID_FOREGROUND);
		font7.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font7.setColor(HSSFColor.WHITE.index);
		style7.setFont(font);
		style7.setAlignment(CellStyle.ALIGN_CENTER);


		studentSheet.addMergedRegion(new CellRangeAddress(7, 7, 0, 10));
		HSSFRow aRow7 = studentSheet.createRow(7);

		aRow7.createCell(0).setCellValue("All Business Areas");
		aRow7.getCell(0).setCellStyle(style7);
		aRow7.setHeight((short) 400);

		studentSheet.addMergedRegion(new CellRangeAddress(8, 9, 0, 0));
		studentSheet.addMergedRegion(new CellRangeAddress(8, 8, 2, 4));
		studentSheet.addMergedRegion(new CellRangeAddress(8, 8, 5, 7));
		studentSheet.addMergedRegion(new CellRangeAddress(8, 8, 8, 10));

		// create style for Row 8
		CellStyle style8 = workbook.createCellStyle();
		Font font8 = workbook.createFont();
		font8.setFontName("Arial");
		style8.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
		style8.setFillPattern(CellStyle.SOLID_FOREGROUND);
		font8.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font8.setColor(HSSFColor.WHITE.index);
		style8.setFont(font);
		style8.setAlignment(CellStyle.ALIGN_CENTER);

		HSSFRow aRow8 = studentSheet.createRow(8);

		aRow8.createCell(0).setCellValue("Year Created");
		aRow8.getCell(0).setCellStyle(style8);

		aRow8.createCell(1).setCellValue("# ideas Submitted");
		aRow8.getCell(1).setCellStyle(style8);

		aRow8.createCell(2).setCellValue("# ideas Resolved in 2014");
		aRow8.getCell(2).setCellStyle(style8);

		aRow8.createCell(5).setCellValue("# ideas Resolved in 2015");
		aRow8.getCell(5).setCellStyle(style8);

		aRow8.createCell(8).setCellValue("# ideas Resolved in 2016");
		aRow8.getCell(8).setCellStyle(style8);

		studentSheet.addMergedRegion(new CellRangeAddress(8, 9, 1, 1));

		HSSFRow aRow9 = studentSheet.createRow(9);
		aRow9.createCell(2).setCellValue("Solved in 2014");
		aRow9.getCell(2).setCellStyle(style6);
		aRow9.setHeight((short) 500);

		aRow9.createCell(3).setCellValue("Closed in 2014");
		aRow9.getCell(3).setCellStyle(style6);
		aRow9.setHeight((short) 500);

		aRow9.createCell(4).setCellValue("Duplicate in 2014");
		aRow9.getCell(4).setCellStyle(style6);
		aRow9.setHeight((short) 500);

		aRow9.createCell(5).setCellValue("Solved in 2015");
		aRow9.getCell(5).setCellStyle(style6);
		aRow9.setHeight((short) 500);

		aRow9.createCell(6).setCellValue("Closed in 2015");
		aRow9.getCell(6).setCellStyle(style6);
		aRow9.setHeight((short) 500);

		aRow9.createCell(7).setCellValue("Duplicate in 2015");
		aRow9.getCell(7).setCellStyle(style6);
		aRow9.setHeight((short) 500);

		aRow9.createCell(8).setCellValue("Solved in 2016");
		aRow9.getCell(8).setCellStyle(style6);
		aRow9.setHeight((short) 500);

		aRow9.createCell(9).setCellValue("Closed in 2016");
		aRow9.getCell(9).setCellStyle(style6);
		aRow9.setHeight((short) 500);

		aRow9.createCell(10).setCellValue("Duplicate in 2016");
		aRow9.getCell(10).setCellStyle(style6);
		aRow9.setHeight((short) 500);




		// create data rows for Student
		int rowCountStudent = 10;

		for (Student aStudent : studentList) {
			HSSFRow aRoww = studentSheet.createRow(rowCountStudent++);
			aRoww.createCell(0).setCellValue(aStudent.getYearCreated());
			aRoww.createCell(1).setCellValue(aStudent.getIdeasSubmitted());
			aRoww.createCell(2).setCellValue(aStudent.getSolvedIn2014A());
			aRoww.createCell(3).setCellValue(aStudent.getClosededIn2014A());
			aRoww.createCell(4).setCellValue(aStudent.getDuplicateIn2014A());
			aRoww.createCell(5).setCellValue(aStudent.getSolvedIn2014B());
			aRoww.createCell(6).setCellValue(aStudent.getClosededIn2014B());
			aRoww.createCell(7).setCellValue(aStudent.getDuplicateIn2014B());
			aRoww.createCell(8).setCellValue(aStudent.getSolvedIn2014C());
			aRoww.createCell(9).setCellValue(aStudent.getClosededIn2014C());
			aRoww.createCell(10).setCellValue(aStudent.getDuplicateIn2014C());
		}
		String latestColumn=String.valueOf(rowCountStudent);

		String cellsFrom_B10_to_B13 = "B10:B"+latestColumn;
		String cellsFrom_C10_to_C13 = "C10:C"+latestColumn;
		String cellsFrom_D10_to_D13 = "D10:D"+latestColumn;
		String cellsFrom_E10_to_E13 = "E10:E"+latestColumn;
		String cellsFrom_F10_to_F13 = "F10:F"+latestColumn;
		String cellsFrom_G10_to_G13 = "G10:G"+latestColumn;
		String cellsFrom_H10_to_H13 = "H10:H"+latestColumn;
		String cellsFrom_I10_to_I13 = "I10:I"+latestColumn;
		String cellsFrom_J10_to_J13 = "J10:J"+latestColumn;
		String cellsFrom_K10_to_K13 = "K10:K"+latestColumn;


		System.out.println(cellsFrom_B10_to_B13);
		HSSFRow aRowN = studentSheet.createRow(rowCountStudent++);
		aRowN.createCell(0).setCellValue("Total");
		//aRowN.createCell(1).setCellFormula("SUM(B10:B13)");
		aRowN.createCell(1).setCellFormula("SUM("+cellsFrom_B10_to_B13+")");
		aRowN.createCell(2).setCellFormula("SUM("+cellsFrom_C10_to_C13+")");
		aRowN.createCell(3).setCellFormula("SUM("+cellsFrom_D10_to_D13+")");
		aRowN.createCell(4).setCellFormula("SUM("+cellsFrom_E10_to_E13+")");
		aRowN.createCell(5).setCellFormula("SUM("+cellsFrom_F10_to_F13+")");
		aRowN.createCell(6).setCellFormula("SUM("+cellsFrom_G10_to_G13+")");
		aRowN.createCell(7).setCellFormula("SUM("+cellsFrom_H10_to_H13+")");
		aRowN.createCell(8).setCellFormula("SUM("+cellsFrom_I10_to_I13+")");
		aRowN.createCell(9).setCellFormula("SUM("+cellsFrom_J10_to_J13+")");	        
		aRowN.createCell(10).setCellFormula("SUM("+cellsFrom_K10_to_K13+")");

		//FOR SECOND PART***********************************************************************
		studentSheet.addMergedRegion(new CellRangeAddress(14, 14, 0, 10));
		HSSFRow aRow14 = studentSheet.createRow(14);

		aRow14.createCell(0).setCellValue("Service Delivery");
		aRow14.getCell(0).setCellStyle(style7);
		aRow14.setHeight((short) 400);

		studentSheet.addMergedRegion(new CellRangeAddress(15, 16, 0, 0));
		studentSheet.addMergedRegion(new CellRangeAddress(15, 16, 1, 1));
		studentSheet.addMergedRegion(new CellRangeAddress(15, 15, 2, 4));
		studentSheet.addMergedRegion(new CellRangeAddress(15, 15, 5, 7));
		studentSheet.addMergedRegion(new CellRangeAddress(15, 15, 8, 10));

		HSSFRow aRow15 = studentSheet.createRow(15);

		aRow15.createCell(0).setCellValue("Year Created");
		aRow15.getCell(0).setCellStyle(style8);

		aRow15.createCell(1).setCellValue("# ideas Submitted");
		aRow15.getCell(1).setCellStyle(style8);

		aRow15.createCell(2).setCellValue("# ideas Resolved in 2014");
		aRow15.getCell(2).setCellStyle(style8);

		aRow15.createCell(5).setCellValue("# ideas Resolved in 2015");
		aRow15.getCell(5).setCellStyle(style8);

		aRow15.createCell(8).setCellValue("# ideas Resolved in 2016");
		aRow15.getCell(8).setCellStyle(style8);

		HSSFRow aRow16 = studentSheet.createRow(16);
		aRow16.createCell(2).setCellValue("Solved in 2014");
		aRow16.getCell(2).setCellStyle(style6);
		aRow16.setHeight((short) 500);

		aRow16.createCell(3).setCellValue("Closed in 2014");
		aRow16.getCell(3).setCellStyle(style6);
		aRow16.setHeight((short) 500);

		aRow16.createCell(4).setCellValue("Duplicate in 2014");
		aRow16.getCell(4).setCellStyle(style6);
		aRow16.setHeight((short) 500);

		aRow16.createCell(5).setCellValue("Solved in 2015");
		aRow16.getCell(5).setCellStyle(style6);
		aRow16.setHeight((short) 500);

		aRow16.createCell(6).setCellValue("Closed in 2015");
		aRow16.getCell(6).setCellStyle(style6);
		aRow16.setHeight((short) 500);

		aRow16.createCell(7).setCellValue("Duplicate in 2015");
		aRow16.getCell(7).setCellStyle(style6);
		aRow16.setHeight((short) 500);

		aRow16.createCell(8).setCellValue("Solved in 2016");
		aRow16.getCell(8).setCellStyle(style6);
		aRow16.setHeight((short) 500);

		aRow16.createCell(9).setCellValue("Closed in 2016");
		aRow16.getCell(9).setCellStyle(style6);
		aRow16.setHeight((short) 500);

		aRow16.createCell(10).setCellValue("Duplicate in 2016");
		aRow16.getCell(10).setCellStyle(style6);
		aRow16.setHeight((short) 500);

		// create data rows for Student
		int rowCountStudent1 = 17;

		for (Student aStudent : studentList) {
			HSSFRow aRoww = studentSheet.createRow(rowCountStudent1++);
			aRoww.createCell(0).setCellValue(aStudent.getYearCreated());
			aRoww.createCell(1).setCellValue(aStudent.getIdeasSubmitted());
			aRoww.createCell(2).setCellValue(aStudent.getSolvedIn2014A());
			aRoww.createCell(3).setCellValue(aStudent.getClosededIn2014A());
			aRoww.createCell(4).setCellValue(aStudent.getDuplicateIn2014A());
			aRoww.createCell(5).setCellValue(aStudent.getSolvedIn2014B());
			aRoww.createCell(6).setCellValue(aStudent.getClosededIn2014B());
			aRoww.createCell(7).setCellValue(aStudent.getDuplicateIn2014B());
			aRoww.createCell(8).setCellValue(aStudent.getSolvedIn2014C());
			aRoww.createCell(9).setCellValue(aStudent.getClosededIn2014C());
			aRoww.createCell(10).setCellValue(aStudent.getDuplicateIn2014C());
		}
		String latestColumn1=String.valueOf(rowCountStudent1);

		String cellsFrom_B17_to_B20 = "B17:B"+latestColumn1;
		String cellsFrom_C17_to_C20 = "C17:C"+latestColumn1;
		String cellsFrom_D17_to_D20 = "D17:D"+latestColumn1;
		String cellsFrom_E17_to_E20 = "E17:E"+latestColumn1;
		String cellsFrom_F17_to_F20 = "F17:F"+latestColumn1;
		String cellsFrom_G17_to_G20 = "G17:G"+latestColumn1;
		String cellsFrom_H17_to_H20 = "H17:H"+latestColumn1;
		String cellsFrom_I17_to_I20 = "I17:I"+latestColumn1;
		String cellsFrom_J17_to_J20 = "J17:J"+latestColumn1;
		String cellsFrom_K17_to_K20 = "K17:K"+latestColumn1;


		System.out.println(cellsFrom_B17_to_B20);
		HSSFRow aRowM = studentSheet.createRow(rowCountStudent1++);
		aRowM.createCell(0).setCellValue("Total");
		//aRowM.createCell(1).setCellFormula("SUM(B17:B20)");
		aRowM.createCell(1).setCellFormula("SUM("+cellsFrom_B17_to_B20+")");
		aRowM.createCell(2).setCellFormula("SUM("+cellsFrom_C17_to_C20+")");
		aRowM.createCell(3).setCellFormula("SUM("+cellsFrom_D17_to_D20+")");
		aRowM.createCell(4).setCellFormula("SUM("+cellsFrom_E17_to_E20+")");
		aRowM.createCell(5).setCellFormula("SUM("+cellsFrom_F17_to_F20+")");
		aRowM.createCell(6).setCellFormula("SUM("+cellsFrom_G17_to_G20+")");
		aRowM.createCell(7).setCellFormula("SUM("+cellsFrom_H17_to_H20+")");
		aRowM.createCell(8).setCellFormula("SUM("+cellsFrom_I17_to_I20+")");
		aRowM.createCell(9).setCellFormula("SUM("+cellsFrom_J17_to_J20+")");	        
		aRowM.createCell(10).setCellFormula("SUM("+cellsFrom_K17_to_K20+")");

	}



}