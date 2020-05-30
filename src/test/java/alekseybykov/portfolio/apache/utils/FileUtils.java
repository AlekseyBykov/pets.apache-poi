package alekseybykov.portfolio.apache.utils;

import java.io.File;
import java.util.Objects;

/**
 * @author Aleksey Bykov
 * @since 30.05.2020
 */
public class FileUtils {
	public static File getFileByName(String fileName) {
		return new File(Objects.requireNonNull(FileUtils.class.getClassLoader().getResource(fileName)).getFile());
	}
}
