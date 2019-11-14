package exceltools.com.excel.handler;

import exceltools.com.excel.base.ExcelUtils;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * @author sunc
 * @date 2019/11/13 19:14
 * @description MergeHandler
 */

public class MergeHandler {

    public static void merge(String destination, List<File> files) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        ExcelUtils.writeExcelCell(destination, files);
    }

}
