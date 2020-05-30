package alekseybykov.portfolio.apache.poiooxml;

import alekseybykov.portfolio.apache.poiooxml.model.Book;
import org.apache.commons.io.FileUtils;
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
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.BeforeClass;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.logging.Logger;

/**
 * @author Aleksey Bykov
 * @since 28.05.2020
 */
public class ApachePOIOOXMLTestBase {

	protected static final Logger logger = Logger.getLogger(ApachePOIOOXMLTestBase.class.getPackage().getName());

	protected final List<String> datesFixture = Arrays.asList(
		"28/05/20", "27/05/20", "26/05/20", "25/05/20",
		"24/05/20", "23/05/20", "22/05/20", "21/05/20",
		"20/05/20", "19/05/20", "18/05/20", "17/05/20",
		"16/05/20", "15/05/20", "14/05/20", "13/05/20",
		"12/05/20", "11/05/20", "10/05/20", "09/05/20",
		"08/05/20", "07/05/20", "06/05/20", "05/05/20",
		"04/05/20", "03/05/20", "02/05/20", "01/05/20"
	);

	protected final List<String> eursFixture = Arrays.asList(
		"0.9078", "0.9098", "0.9112", "0.9166", "0.9171",
		"0.9171", "0.9171", "0.9091", "0.9126", "0.9132",
		"0.9232", "0.9261", "0.9261", "0.9261", "0.9266",
		"0.9195", "0.921", "0.9239", "0.9223", "0.9223",
		"0.9223", "0.9274", "0.9253", "0.9223", "0.9139",
		"0.9195", "0.9195", "0.9195"
	);

	protected final List<String> gbpsFixture = Arrays.asList(
			"0.8145", "0.8152", "0.8098", "0.8205", "0.8214",
			"0.8214", "0.8214", "0.8177", "0.8155", "0.8177",
			"0.8231", "0.8218", "0.8218", "0.8218", "0.82",
			"0.8114", "0.8084", "0.8119", "0.8073", "0.8073",
			"0.8073", "0.8113", "0.8074", "0.8029", "0.8033",
			"0.7991", "0.7991", "0.7991"
	);

	@BeforeClass
	public static void init() throws IOException {
		System.setProperty("java.util.logging.config.file", ClassLoader.getSystemResource("logging.properties").getPath());
		FileUtils.cleanDirectory(new File("xlsx"));
	}

	protected void createAndPopulateXlsxWithData(String path, List<Book> books) {
		try (Workbook workbook = createXlsxWorkbook(books);
		     FileOutputStream outputStream = new FileOutputStream(path)) {
			workbook.write(outputStream);
		} catch (IOException ex) {
			logger.severe("Error while reading the file:" + ex.getMessage());
		}
	}

	private Workbook createXlsxWorkbook(List<Book> books) {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("TAB NAME");
		createHeaderArea(sheet);
		createTableArea(sheet, books);
		setAutoFilter(sheet);
		setZoom(sheet);
		return workbook;
	}

	private void createHeaderArea(Sheet sheet) {
		Row row = sheet.createRow(0);
		CellStyle cellStyle = createStyleForHeaderCell(sheet);
		for (int i = 0; i < 5; i++) {
			Cell cell = row.createCell(i);
			sheet.autoSizeColumn(i);

			if (i == 0) {
				cell.setCellValue("ISBN");
				cell.setCellStyle(cellStyle);
			} else if (i == 1) {
				cell.setCellValue("BOOK TITLE");
				cell.setCellStyle(cellStyle);
			} else if (i == 2) {
				cell.setCellValue("AUTHOR");
				cell.setCellStyle(cellStyle);
			} else if (i == 3) {
				cell.setCellValue("PUBLISHER");
				cell.setCellStyle(cellStyle);
			} else {
				cell.setCellValue("PRICE");
				cell.setCellStyle(cellStyle);
			}
		}
	}

	private void createTableArea(Sheet sheet, List<Book> books) {
		CellStyle cellStyle = createStyleForTableCell(sheet);
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

	private CellStyle createStyleForHeaderCell(Sheet sheet) {
		Font font = sheet.getWorkbook().createFont();
		font.setBold(true);
		font.setFontName("Arial");
		font.setFontHeightInPoints((short)21);
		font.setColor(IndexedColors.WHITE.getIndex());

		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setFont(font);

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

	private CellStyle createStyleForTableCell(Sheet sheet) {
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

	private void setAutoFilter(Sheet sheet) {
		sheet.setAutoFilter(new CellRangeAddress(sheet.getFirstRowNum(), sheet.getLastRowNum(),
				sheet.getRow(0).getFirstCellNum(), sheet.getRow(0).getLastCellNum() - 1));
	}

	private void setZoom(Sheet sheet) {
		sheet.setZoom(150);
	}
}
