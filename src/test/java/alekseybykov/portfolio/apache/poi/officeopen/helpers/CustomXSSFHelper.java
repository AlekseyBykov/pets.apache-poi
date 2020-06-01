package alekseybykov.portfolio.apache.poi.officeopen.helpers;

import alekseybykov.portfolio.apache.poi.officeopen.model.Book;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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
	private final static int isbnIdx = 0;
	private final static int titleIdx = 1;
	private final static int authorIdx = 2;
	private final static int publisherIdx = 3;

	public static void createHeaderAreaForSheet(Sheet sheet) {
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
				if (j == isbnIdx) {
					cell.setCellValue(books.get(i).getIsbn());
				} else if (j == titleIdx) {
					cell.setCellValue(books.get(i).getTitle());
				} else if (j == authorIdx) {
					cell.setCellValue(books.get(i).getAuthor());
				} else if (j == publisherIdx) {
					cell.setCellValue(books.get(i).getPublisher());
				} else {
					cell.setCellValue(books.get(i).getPrice());
				}
				cell.setCellStyle(cellStyle);
			}
		}
	}

	public static void setAutoFilterForSheet(Sheet sheet) {
		sheet.setAutoFilter(new CellRangeAddress(sheet.getFirstRowNum(), sheet.getLastRowNum(),
				sheet.getRow(0).getFirstCellNum(), sheet.getRow(0).getLastCellNum() - 1));
	}

	public static void setZoomForSheet(Sheet sheet) {
		sheet.setZoom(110);
	}

	private static CellStyle createStyleForTableCellInSheet(Sheet sheet) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setFont(createFont(sheet, (short) 10));

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
		cellStyle.setFont(createFont(sheet, (short) 16));

		cellStyle.setAlignment(HorizontalAlignment.LEFT);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);

		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);

		return cellStyle;
	}

	private static Font createFont(Sheet sheet, short fontSize) {
		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints(fontSize);
		font.setColor(IndexedColors.BLACK.getIndex());
		return font;
	}
}
