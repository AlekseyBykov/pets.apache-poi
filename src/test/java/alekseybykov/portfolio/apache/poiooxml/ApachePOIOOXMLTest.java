package alekseybykov.portfolio.apache.poiooxml;

import alekseybykov.portfolio.apache.poiooxml.model.Book;
import alekseybykov.portfolio.apache.poiooxml.model.BookBuilder;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;

import static java.lang.String.format;
import static org.junit.Assert.assertEquals;

// For MS Excel 2007
public class ApachePOIOOXMLTest extends ApachePOIOOXMLTestBase {
	private static final Logger LOGGER = Logger.getLogger(ApachePOIOOXMLTest.class.getPackage().getName());

	@Test
	public void testCreateEmptyXlsxFile() {
		try (Workbook workbook = new XSSFWorkbook();
		        FileOutputStream outputStream = new FileOutputStream("xlsx/empty.xlsx")) {
			workbook.createSheet("TAB_NAME");
			workbook.write(outputStream);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	@Test
	public void testCreateXlsxFileWithData() {
		List<Book> books = new ArrayList<>();
		for (int i = 0; i < 100; i ++) {
			books.add(new BookBuilder()
				.setIsbn(format("%s%d", "xxx-xxx-xxx-", i))
				.setAuthor(format("%s%d", "Author ", i))
				.setPrice(1.1f)
				.setPublisher(format("%s%d", "Publisher ", i))
				.setTitle(format("%s%d", "Title ", i))
				.build());
		}

		createAndPopulateXlsxWithData("xlsx/with-data.xlsx", books);
	}

	@Test
	public void testReadXlsx() {
		DataFormatter dataFormatter = new DataFormatter();
		try (InputStream inputStream = new FileInputStream(getFile("exchange-rates-2020-may.xlsx"))) {
			Workbook workbook = WorkbookFactory.create(inputStream);
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals("Exchange rates", sheet.getSheetName());
			for (int rowNum = sheet.getFirstRowNum(); rowNum <= sheet.getLastRowNum(); rowNum++) {
				Row row = sheet.getRow(rowNum);
				for (int columnNum = 0; columnNum < row.getLastCellNum(); columnNum++) {
					Cell cell = row.getCell(columnNum);
					if (cell != null) {
						String value = dataFormatter.formatCellValue(cell);
						// todo: simplification
						if (rowNum == 1) {
							if (columnNum == 0) {
								assertEquals("28/05/20", value);
							} else if (columnNum == 1) {
								assertEquals("1", value);
							} else if (columnNum == 2) {
								assertEquals("0,9078", value);
							} else if (columnNum == 3) {
								assertEquals("0,8145", value);
							}
						}
					}
				}
			}

		} catch (IOException ex) {
			LOGGER.severe("Error while reading the file:" + ex.getMessage());
		} catch (InvalidFormatException ex) {
			LOGGER.severe("Invalid file format:" + ex.getMessage());
		}
	}
}
