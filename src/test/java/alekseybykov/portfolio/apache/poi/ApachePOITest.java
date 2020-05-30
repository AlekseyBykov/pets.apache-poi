package alekseybykov.portfolio.apache.poi;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Examples of manipulations with MS Excel 97.
 *
 * @author Aleksey Bykov
 * @since 28.05.2020
 */
public class ApachePOITest {

	@Before
	public void init() throws IOException {
		FileUtils.cleanDirectory(new File("xls"));
	}

	@Test
	public void testCreateEmptyXlsFile() {
		try (Workbook workbook = new HSSFWorkbook();
		        FileOutputStream fileOutputStream = new FileOutputStream("xls/empty.xls")) {
			workbook.createSheet("TAB_NAME");
			workbook.write(fileOutputStream);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
