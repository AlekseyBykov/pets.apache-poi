package alekseybykov.portfolio.apache.poiooxml;

import alekseybykov.portfolio.apache.poiooxml.consts.Consts;
import alekseybykov.portfolio.apache.poiooxml.model.Book;
import alekseybykov.portfolio.apache.poiooxml.model.BookBuilder;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static java.lang.String.format;

public class ApachePOIOOXMLTest extends ApachePOIOOXMLTestBase {

	@Test
	public void testCreateEmptyXlsxFile() {
		try (Workbook workbook = new XSSFWorkbook();
		        FileOutputStream outputStream = new FileOutputStream("xlsx/empty.xlsx")) {
			workbook.createSheet(Consts.TAB_NAME.toString());
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
}
