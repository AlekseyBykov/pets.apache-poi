package alekseybykov.portfolio.apache.poi;

import alekseybykov.portfolio.apache.poiooxml.consts.Consts;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import static alekseybykov.portfolio.apache.poiooxml.consts.Consts.TAB_NAME;

public class ApachePOITest {

	@Before
	public void init() throws IOException {
		FileUtils.cleanDirectory(new File("xls"));
	}

	@Test
	public void testCreateEmptyXlsFile() {
		try (Workbook workbook = new HSSFWorkbook();
		        FileOutputStream fileOutputStream = new FileOutputStream("xls/empty.xls")) {
			workbook.createSheet(TAB_NAME.toString());
			workbook.write(fileOutputStream);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
