package alekseybykov.portfolio.apachepoi.poiooxml;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

public class ApachePOIOOXMLTest {

	@Test
	public void testGenerateModernFormatXlsxFile() {
		Workbook workbook = new XSSFWorkbook();
		workbook.createSheet("Tab name here");
		try (FileOutputStream outputStream = new FileOutputStream("xlsx/empty.xlsx")) {
			workbook.write(outputStream);
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
