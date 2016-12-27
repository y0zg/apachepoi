package com.parsefile;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.String;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.lang.Object;

public class ParseFileXLS  {

	public static ArrayList GetColumn() throws IOException {
		ArrayList list = new ArrayList();

		File excel = new File("C:\\temp\\file1.xlsx");
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		XSSFSheet sheet = book.getSheet("Sheet1");

		for (Row row : sheet) {
			for (Cell cell : row) {
				list.add(cell.getStringCellValue());
			}
			break;
		}
		return list;
	}

	public static ArrayList AddColumn(ArrayList args) throws IOException {
		ArrayList columns = args;

		File excel = new File("C:\\temp\\file1.xlsx");
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		XSSFSheet sheet = book.getSheet("Sheet1");
		Row header = sheet.getRow(0);
		Integer colsize = GetColumn().size();
		Integer colsum = colsize + columns.size();

		for (int i = colsize; i < colsum; i++) {
			String col = columns.get(i - colsize).toString();
			header.createCell(i).setCellValue(col);
		}

		File excelOut = new File("C:\\temp\\fileOut1.xlsx");
		FileOutputStream fileOut = new FileOutputStream(excelOut);
		book.write(fileOut);
		fileOut.close();

		return columns;
	}


	public static ArrayList DelColumn(ArrayList args) throws IOException {
		ArrayList columns = args;

		File excel = new File("C:\\temp\\file1.xlsx");
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		XSSFSheet sheet = book.getSheet("Sheet1");

		for (int a = 0; a < sheet.getPhysicalNumberOfRows(); a++) {
			Row header = sheet.getRow(a);

			for (int i = 0; i < columns.size(); i++) {
				String col = columns.get(i).toString();
				Cell oldcell = header.getCell(i);
				header.removeCell(oldcell);
			}
		}

		File excelOut = new File("C:\\temp\\fileOut1.xlsx");
		FileOutputStream fileOut = new FileOutputStream(excelOut);
		book.write(fileOut);
		fileOut.close();

		return columns;
	}

	public static ArrayList ModifyColumn(ArrayList args) throws IOException {
		ArrayList columns = args;
		ArrayList allcol = GetColumn();

		File excel = new File("C:\\temp\\file1.xlsx");
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		XSSFSheet sheet = book.getSheet("Sheet1");

		for (int a = 0; a < sheet.getPhysicalNumberOfRows(); a++) {
			Row header = sheet.getRow(a);

			for (int i = 0; i < allcol.size(); i++) {

				if (!columns.contains(allcol.get(i))) {
					Cell oldcell = header.getCell(i);
					header.removeCell(oldcell);
				}
			}
		}

		File excelOut = new File("C:\\temp\\fileOut1.xlsx");
		FileOutputStream fileOut = new FileOutputStream(excelOut);
		book.write(fileOut);
		fileOut.close();

		return columns;
	}

	public static void main(String[] args) throws IOException {

		if(args[0].equals("get") == true) {
			ArrayList result = GetColumn();
			System.out.println(result);
		}

		if(args[0].equals("modify") == true) {
			ArrayList param = new ArrayList(Arrays.asList("col1","col3"));
			ArrayList result = ModifyColumn(param);
			System.out.println(result);
		}

		if(args[0].equals("add") == true) {
			ArrayList param = new ArrayList(Arrays.asList(args[1].split(",")));
			ArrayList result = AddColumn(param);
			System.out.println(result);
		}

		if(args[0].equals("del") == true) {
			ArrayList param = new ArrayList(Arrays.asList(args[1].split(",")));
			ArrayList result = DelColumn(param);
			System.out.println(result);
		}

		if(args.length == 0) {
			System.out.println("No argument");
			System.exit(0);
		}

	}

}

