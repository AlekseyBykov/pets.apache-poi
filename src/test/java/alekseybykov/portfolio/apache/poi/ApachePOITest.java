package alekseybykov.portfolio.apache.poi;

import alekseybykov.portfolio.apache.poi.officeopen.helpers.CustomFileHelper;
import alekseybykov.portfolio.apache.poi.officeopen.model.Book;
import alekseybykov.portfolio.apache.poi.officeopen.model.BookBuilder;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
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
import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertThat;

/**
 * Examples of manipulations with MS Excel 97 and 2007.
 *
 * @author Aleksey Bykov
 * @since 28.05.2020
 */
public class ApachePOITest extends ApachePOITestBase {

	@Test
	public void testCreateEmptyXlsFile() {
		try (Workbook workbook = new HSSFWorkbook();
		        FileOutputStream fileOutputStream = new FileOutputStream("xls/empty.xls")) {
			workbook.createSheet("TAB_NAME");
			workbook.write(fileOutputStream);
		} catch (IOException ex) {
			logger.severe("Error while writing to the file: " + ex.getMessage());
		}
	}

	@Test
	public void testCreateEmptyXlsxWithOneSheet() {
		try (Workbook workbook = new XSSFWorkbook();
		     FileOutputStream outputStream = new FileOutputStream("xlsx/empty.xlsx")) {
			workbook.createSheet("TAB_NAME");
			workbook.write(outputStream);
		} catch (IOException ex) {
			logger.severe("Error while writing to the file: " + ex.getMessage());
		}
	}

	@Test
	public void testCreateXlsxFileWithOneSheetAndData() {
		List<Book> books = new ArrayList<>();
		for (int i = 0; i < 100; i ++) {
			books.add(new BookBuilder()
					.setIsbn(format("%s%d", "xxx-xxx-xxx-", i))
					.setAuthor(format("%s%d", "Author ", i))
					.setPrice(i)
					.setPublisher(format("%s%d", "Publisher ", i))
					.setTitle(format("%s%d", "Title ", i))
					.build());
		}

		createAndPopulateXlsxWithData("xlsx/with-data.xlsx", books);
	}

	@Test
	public void testReadXlsxWithOneSheet() {
		try (InputStream inputStream = new FileInputStream(CustomFileHelper.getFileByName("exchange-rates-2020-may.xlsx"))) {
			List<String> dates = new ArrayList<>();
			List<String> eurs = new ArrayList<>();
			List<String> gbps = new ArrayList<>();

			final int dateIdx = 0;
			final int eurIdx = 2;
			final int gbpIdx = 3;
			final int hdrOffset = 1;

			DataFormatter dataFormatter = new DataFormatter();
			Workbook workbook = WorkbookFactory.create(inputStream);
			Sheet sheet = workbook.getSheetAt(NumberUtils.INTEGER_ZERO);
			assertEquals("exchange rates", sheet.getSheetName());
			for (int rowNum = sheet.getFirstRowNum() + hdrOffset; rowNum <= sheet.getLastRowNum(); rowNum++) {
				Row row = sheet.getRow(rowNum);
				for (int columnNum = NumberUtils.INTEGER_ZERO; columnNum < row.getLastCellNum(); columnNum++) {
					Cell cell = row.getCell(columnNum);
					if (cell != null) {
						String value = dataFormatter.formatCellValue(cell);
						if (columnNum == dateIdx) {
							dates.add(value);
						} else if (columnNum == eurIdx) {
							eurs.add(value);
						} else if (columnNum == gbpIdx) {
							gbps.add(value);
						}
					}
				}
			}

			assertThat(dates, is(datesFixture));
			assertThat(eurs, is(eursFixture));
			assertThat(gbps, is(gbpsFixture));

		} catch (IOException ex) {
			logger.severe("Error while reading the file: " + ex.getMessage());
		} catch (InvalidFormatException ex) {
			logger.severe("Invalid file format: " + ex.getMessage());
		}
	}
}
