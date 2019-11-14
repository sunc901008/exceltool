package exceltools.com.excel.base;

import java.util.Arrays;
import java.util.List;

public class Constant {

    static final String DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";

    static final List<String> TYPES = Arrays.asList("xlsx", "xls");

    public static String BASE_PATH = System.getProperty("user.dir").substring(0, 3);


}
