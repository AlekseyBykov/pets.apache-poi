package alekseybykov.portfolio.apachepoi.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

public class ApachePOITest {

	@Test
	public void testGenerateOldFormatXlsFile() {
		Workbook workbook = new HSSFWorkbook();
		workbook.createSheet("Tab name here");
		try (FileOutputStream outputStream = new FileOutputStream("xls/simple.xls")) {
			workbook.write(outputStream);
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
