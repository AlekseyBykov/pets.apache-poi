package alekseybykov.portfolio.apache.poiooxml.helpers;

import java.io.File;
import java.util.Objects;

/**
 * @author Aleksey Bykov
 * @since 30.05.2020
 */
public class CustomFileHelper {
	public static File getFileByName(String fileName) {
		return new File(Objects.requireNonNull(CustomFileHelper.class.getClassLoader().getResource(fileName)).getFile());
	}
}
