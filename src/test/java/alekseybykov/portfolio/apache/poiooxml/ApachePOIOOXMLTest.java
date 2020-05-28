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
import java.util.Arrays;
import java.util.List;

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
		Book book1 = new BookBuilder()
				.setIsbn("xxx-xxx-xxx-1")
				.setAuthor("Author 1")
				.setPrice(1.1f)
				.setPublisher("Publisher 1")
				.setTitle("Title 1")
				.build();

		Book book2 = new BookBuilder()
				.setIsbn("xxx-xxx-xxx-2")
				.setAuthor("Author 2")
				.setPrice(1.2f)
				.setPublisher("Publisher 2")
				.setTitle("Title 2")
				.build();

		List<Book> books = new ArrayList<>(Arrays.asList(book1, book2));
		createAndPopulateXlsxWithData("xlsx/with-data.xlsx", books);
	}
}
