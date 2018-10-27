package com.excel.poi;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SuppressWarnings("all")
public class ExcelExporter {

	/**
	 * 获取表头的样式
	 * @param wb Workbook
	 * @return CellStyle
	 */
	private static CellStyle headerStyle(Workbook wb) {
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);

		cellStyle.setAlignment(HorizontalAlignment.LEFT);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

		//设置表格的颜色
//		cellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
//		cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		return cellStyle;
	}

	/**
	 * 获取数据的样式
	 * @param wb Workbook
	 * @return CellStyle
	 */
	private static CellStyle dataStyle(Workbook wb) {
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);

		cellStyle.setAlignment(HorizontalAlignment.LEFT);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		return cellStyle;
	}

	/**
	 * 获取日期列样式
	 * @param wb Workbook
	 * @return 日期列样式
	 */
	private static CellStyle getDateCellType(Workbook wb) {
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setDataFormat(wb.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);

		cellStyle.setAlignment(HorizontalAlignment.LEFT);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		return cellStyle;
	}

	/**
	 * 用行列数据生成简单的Excel
	 * 
	 * @param columns
	 * @param rows
	 * @throws IOException
	 */
	public static void simpleExcel(List<String> columns, List<Map<String, Object>> rows, OutputStream outputStream,
			ProgressReporter progress) throws IOException {
		if (progress == null) {
			progress = NullProgressReporter.instance;
		}

		Workbook wb = new XSSFWorkbook();

		Sheet createSheet = wb.createSheet();

		//冻结首行
		createSheet.createFreezePane(0,1);

		int startRow = 0;
		int startCol = 0;

		// 写入Header
		org.apache.poi.ss.usermodel.Row headerRow = createSheet.createRow(startRow);
		for (int colNum = 0; colNum < columns.size(); ++colNum) {
			org.apache.poi.ss.usermodel.Cell cell = headerRow.createCell(startCol + colNum);
			cell.setCellValue(columns.get(colNum));
			cell.setCellStyle(headerStyle(wb));
		}

		// 写入数据
		int dataRowStart = startRow + 1;
		for (int rowNum = 0; rowNum < rows.size(); ++rowNum) {

			progress.setProgress("writeExcel", rowNum * 100/ rows.size(), String.format("写入Excel，%d/%d", rowNum, rows.size()));

			Map<String, Object> row = rows.get(rowNum);

			Row dataRow = createSheet.createRow(dataRowStart + rowNum);


			for (int col = 0; col < columns.size(); ++col) {
				Cell dataCell = dataRow.createCell(startCol + col);
				dataCell.setCellStyle(dataStyle(wb));
				Object value = row.get(columns.get(col));
				if (value == null) {
					// do nothing
				} else if (value instanceof String) {
					dataCell.setCellValue((String) value);
				} else if (value instanceof Date) {
					dataCell.setCellValue((Date) value);
					dataCell.setCellStyle(getDateCellType(wb));
				} else if (TypeUtil.isNumericType(value.getClass())) {
					dataCell.setCellValue(Castors.me().castTo(value, Double.class));
				} else {
					// do nothing
				}
			}
		}

		wb.write(outputStream);
	}
}
