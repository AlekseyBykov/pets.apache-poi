package alekseybykov.portfolio.apache.poiooxml.helpers;

import alekseybykov.portfolio.apache.poiooxml.model.Book;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * @author Aleksey Bykov
 * @since 01.06.2020
 */
public class CustomXSSFHelper {
	public static void createHeaderAreaForSheet(Sheet sheet) {
		final int isbnIdx = 0;
		final int titleIdx = 1;
		final int authorIdx = 2;
		final int publisherIdx = 3;
		Row row = sheet.createRow(0);
		CellStyle cellStyle = createStyleForHeaderInSheet(sheet);
		for (int i = 0; i < 5; i++) {
			Cell cell = row.createCell(i);
			sheet.autoSizeColumn(i);

			if (i == isbnIdx) {
				cell.setCellValue("ISBN");
			} else if (i == titleIdx) {
				cell.setCellValue("BOOK TITLE");
			} else if (i == authorIdx) {
				cell.setCellValue("AUTHOR");
			} else if (i == publisherIdx) {
				cell.setCellValue("PUBLISHER");
			} else {
				cell.setCellValue("PRICE");
			}
			cell.setCellStyle(cellStyle);
		}
	}

	public static void createTableAreaForSheet(Sheet sheet, List<Book> books) {
		CellStyle cellStyle = createStyleForTableCellInSheet(sheet);
		for (int i = 0; i < books.size(); i++) {
			Row row = sheet.createRow(i + 1);
			for (int j = 0; j < 5; j++) {
				Cell cell = row.createCell(j);
				sheet.autoSizeColumn(j);

				if (j == 0) {
					cell.setCellValue(books.get(i).getIsbn());
					cell.setCellStyle(cellStyle);
				} else if (j == 1) {
					cell.setCellValue(books.get(i).getTitle());
					cell.setCellStyle(cellStyle);
				} else if (j == 2) {
					cell.setCellValue(books.get(i).getAuthor());
					cell.setCellStyle(cellStyle);
				} else if (j == 3) {
					cell.setCellValue(books.get(i).getPublisher());
					cell.setCellStyle(cellStyle);
				} else {
					cell.setCellValue(books.get(i).getPrice());
					cell.setCellStyle(cellStyle);
				}
			}
		}
	}

	public static void setAutoFilterForSheet(Sheet sheet) {
		sheet.setAutoFilter(new CellRangeAddress(sheet.getFirstRowNum(), sheet.getLastRowNum(),
				sheet.getRow(0).getFirstCellNum(), sheet.getRow(0).getLastCellNum() - 1));
	}

	public static void setZoomForSheet(Sheet sheet) {
		sheet.setZoom(150);
	}

	private static CellStyle createStyleForTableCellInSheet(Sheet sheet) {
		Font font = sheet.getWorkbook().createFont();
		font.setItalic(true);
		font.setFontName("Arial");
		font.setFontHeightInPoints((short)10);

		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setFont(font);

		cellStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
		cellStyle.setBorderTop(BorderStyle.THIN);

		cellStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
		cellStyle.setBorderRight(BorderStyle.THIN);

		cellStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
		cellStyle.setBorderBottom(BorderStyle.THIN);

		cellStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
		cellStyle.setBorderLeft(BorderStyle.THIN);

		return cellStyle;
	}

	private static CellStyle createStyleForHeaderInSheet(Sheet sheet) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setFont(createFont(sheet));

		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

		cellStyle.setBorderTop(BorderStyle.MEDIUM);
		cellStyle.setBorderBottom(BorderStyle.MEDIUM);

		cellStyle.setBorderRight(BorderStyle.MEDIUM);
		cellStyle.setBorderLeft(BorderStyle.MEDIUM);

		cellStyle.setFillForegroundColor(IndexedColors.GREY_80_PERCENT.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		return cellStyle;
	}

	private static Font createFont(Sheet sheet) {
		Font font = sheet.getWorkbook().createFont();
		font.setBold(true);
		font.setFontName("Arial");
		font.setFontHeightInPoints((short)21);
		font.setColor(IndexedColors.WHITE.getIndex());
		return font;
	}
}
