package exceltools.com.excel.handler;

import exceltools.com.excel.base.ExcelUtils;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;

/**
 * @author sunc
 * @date 2019/11/13 19:14
 * @description MergeHandler
 */

public class SplitHandler {

    public static String split(File file) {
        try {
            return splitExcel(file);
        } catch (OpenXML4JException | ParserConfigurationException | SAXException | IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    private static String splitExcel(File file) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        String destination = null;
        if (file.exists()) {
            if (file.isFile()) {
                destination = file.getAbsolutePath();
                destination = destination.substring(0, destination.lastIndexOf(".")) + String.format("_%s.xls", System.currentTimeMillis());
                ExcelUtils.splitExcelCell(destination, file);
            }
        }
        return destination;
    }

}
