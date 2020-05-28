package alekseybykov.portfolio.apache.poiooxml;

import alekseybykov.portfolio.apache.poiooxml.consts.Consts;
import alekseybykov.portfolio.apache.poiooxml.model.Book;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.BeforeClass;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import static alekseybykov.portfolio.apache.poiooxml.consts.Consts.AUTHOR_TITLE;
import static alekseybykov.portfolio.apache.poiooxml.consts.Consts.BOOK_TITLE;
import static alekseybykov.portfolio.apache.poiooxml.consts.Consts.ISBN_TITLE;
import static alekseybykov.portfolio.apache.poiooxml.consts.Consts.PRICE_TITLE;
import static alekseybykov.portfolio.apache.poiooxml.consts.Consts.PUBLISHER_TITLE;

public class ApachePOIOOXMLTestBase {

	private final int headerRowIdx = 0;
	private final int isbnColumnIdx = 0;
	private final int titleColumnIdx = 1;
	private final int authorColumnIdx = 2;
	private final int publisherColumnIdx = 3;
	private final int priceColumnIdx = 4;

	@BeforeClass
	public static void init() throws IOException {
		FileUtils.cleanDirectory(new File("xlsx"));
	}

	protected void createAndPopulateXlsxWithData(String path, List<Book> books) {
		try (Workbook workbook = createXlsxWorkbook(books);
		     FileOutputStream outputStream = new FileOutputStream(path)) {
			workbook.write(outputStream);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private Workbook createXlsxWorkbook(List<Book> books) {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet(Consts.TAB_NAME.toString());
		createXlsxHeader(sheet);
		createXlsxTable(sheet, books);
		return workbook;
	}

	private void createXlsxHeader(Sheet sheet) {
		Row row = sheet.createRow(headerRowIdx);

		Cell isbnCell = row.createCell(isbnColumnIdx);
		Cell titleCell = row.createCell(titleColumnIdx);
		Cell authorCell = row.createCell(authorColumnIdx);
		Cell publisherCell = row.createCell(publisherColumnIdx);
		Cell priceCell = row.createCell(priceColumnIdx);

		isbnCell.setCellValue(ISBN_TITLE.toString());
		titleCell.setCellValue(BOOK_TITLE.toString());
		authorCell.setCellValue(AUTHOR_TITLE.toString());
		publisherCell.setCellValue(PUBLISHER_TITLE.toString());
		priceCell.setCellValue(PRICE_TITLE.toString());
	}

	private void createXlsxTable(Sheet sheet, List<Book> books) {
		for (int i = 0; i < books.size(); i++) {
			Row row = sheet.createRow(i + 1);

			Cell isbnCell = row.createCell(isbnColumnIdx);
			Cell titleCell = row.createCell(titleColumnIdx);
			Cell authorCell = row.createCell(authorColumnIdx);
			Cell publisherCell = row.createCell(publisherColumnIdx);
			Cell priceCell = row.createCell(priceColumnIdx);

			isbnCell.setCellValue(books.get(i).getIsbn());
			titleCell.setCellValue(books.get(i).getTitle());
			authorCell.setCellValue(books.get(i).getAuthor());
			publisherCell.setCellValue(books.get(i).getPublisher());
			priceCell.setCellValue(books.get(i).getPrice());
		}
	}
}
