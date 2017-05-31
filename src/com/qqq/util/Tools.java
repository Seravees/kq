package com.qqq.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Tools {
	@SuppressWarnings({ "resource", "deprecation" })
	public static List<List<Object>> readAll(String path, String fileName,
			String fileType) throws IOException {
		List<List<Object>> objects = new ArrayList<List<Object>>();

		InputStream stream = new FileInputStream(path + fileName + "."
				+ fileType);
		Workbook wb = null;
		if (fileType.equals("xls")) {
			wb = new HSSFWorkbook(stream);
		} else if (fileType.equals("xlsx")) {
			wb = new XSSFWorkbook(stream);
		} else {
		}
		Sheet sheet = wb.getSheetAt(0);
		// System.out.println(sheet.getLastRowNum());
		for (Row row : sheet) {
			List<Object> temp = new ArrayList<Object>();
			for (Cell cell : row) {
				switch (cell.getCellType()) {
				case HSSFCell.CELL_TYPE_STRING:
					temp.add(cell.getStringCellValue());
					// System.out.println(cell.getStringCellValue());
					break;
				case HSSFCell.CELL_TYPE_NUMERIC:
					temp.add(cell.getNumericCellValue());
					break;
				case HSSFCell.CELL_TYPE_BLANK:
					temp.add(" ");

				default:
					break;
				}
			}
			objects.add(temp);
		}
		return objects;
	}

	@SuppressWarnings("deprecation")
	public static void writerString(String inPath, String inFileName,
			String outPath, String outFileName, String fileType, int rowNum,
			int cellNum, String string, boolean flag) throws IOException {
		File file = new File(outPath);
		file.mkdir();
		InputStream instream = new FileInputStream(inPath + inFileName + "."
				+ fileType);
		Workbook wb = null;
		if (fileType.equals("xls")) {
			wb = new HSSFWorkbook(instream);
		} else if (fileType.equals("xlsx")) {
			wb = new XSSFWorkbook(instream);
		} else {
		}

		Sheet sheet = wb.getSheetAt(0);
		Row row = sheet.getRow(rowNum);
		Cell cell = row.getCell(cellNum);
		if (cell == null) {
			cell = row.createCell(cellNum);
		}

		CellStyle style = wb.createCellStyle();
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setFillForegroundColor(HSSFColor.YELLOW.index);
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);

		CellStyle style2 = wb.createCellStyle();
		style2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style2.setFillForegroundColor(HSSFColor.WHITE.index);
		style2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style2.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style2.setBorderTop(HSSFCellStyle.BORDER_THIN);

		if (flag) {
			if (string.contains("缺勤") || string.contains("否")) {
				cell.setCellStyle(style);
			} else {
				cell.setCellStyle(style2);
			}
		}

		cell.setCellValue(string);

		OutputStream outstream = new FileOutputStream(outPath + outFileName
				+ "." + fileType);

		System.out.println(rowNum + "," + cellNum + " " + string);
		wb.write(outstream);
		// wb.close();
		outstream.close();
	}

	@SuppressWarnings("deprecation")
	public static void writerDouble(String inPath, String inFileName,
			String outPath, String outFileName, String fileType, int rowNum,
			int cellNum, double string, boolean flag) throws IOException {
		InputStream instream = new FileInputStream(inPath + inFileName + "."
				+ fileType);
		Workbook wb = null;
		if (fileType.equals("xls")) {
			wb = new HSSFWorkbook(instream);
		} else if (fileType.equals("xlsx")) {
			wb = new XSSFWorkbook(instream);
		} else {
		}

		Sheet sheet = wb.getSheetAt(0);
		Row row = sheet.getRow(rowNum);
		Cell cell = row.getCell(cellNum);
		if (cell == null) {
			cell = row.createCell(cellNum);
		}
		CellStyle style = wb.createCellStyle();

		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);
		if (flag) {
			cell.setCellStyle(style);
		}
		cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
		cell.setCellValue(string);

		OutputStream outstream = new FileOutputStream(outPath + outFileName
				+ "." + fileType);

		System.out.println(rowNum + "," + cellNum + " " + string);
		wb.write(outstream);
		// wb.close();
		outstream.close();
	}

	public static Workbook open(String path, String fileName, String fileType)
			throws IOException {
		InputStream stream = new FileInputStream(path + fileName + "."
				+ fileType);
		Workbook wb = null;
		if (fileType.equals("xls")) {
			wb = new HSSFWorkbook(stream);
		} else if (fileType.equals("xlsx")) {
			wb = new XSSFWorkbook(stream);
		} else {
		}

		return wb;
	}

	@SuppressWarnings("deprecation")
	public static void shift(String inPath, String inFileName, String outPath,
			String outFileName, String fileType, int startRow)
			throws IOException {
		InputStream instream = new FileInputStream(inPath + inFileName + "."
				+ fileType);
		Workbook wb = null;
		if (fileType.equals("xls")) {
			wb = new HSSFWorkbook(instream);
		} else if (fileType.equals("xlsx")) {
			wb = new XSSFWorkbook(instream);
		} else {
		}

		Sheet sheet = wb.getSheetAt(0);
		sheet.shiftRows(startRow, sheet.getLastRowNum(), 1, true, false);
		Row row = sheet.createRow(startRow);
		row.setHeight(sheet.getRow(startRow - 1).getHeight());
		for (int i = 0; i <= 8; i++) {
			Cell cell = row.createCell(i);
			cell.setCellStyle(sheet.getRow(startRow - 1).getCell(i)
					.getCellStyle());
			switch (sheet.getRow(startRow - 1).getCell(i).getCellType()) {
			case HSSFCell.CELL_TYPE_NUMERIC:
				cell.setCellValue(sheet.getRow(startRow - 1).getCell(i)
						.getNumericCellValue());
				break;
			case HSSFCell.CELL_TYPE_STRING:
				cell.setCellValue(sheet.getRow(startRow - 1).getCell(i)
						.getStringCellValue());
				break;
			default:
				break;
			}
		}

		OutputStream outstream = new FileOutputStream(outPath + outFileName
				+ "." + fileType);
		wb.write(outstream);
		outstream.close();
	}

}
