package exceltools.com.excel.handler;

import exceltools.com.excel.base.Common;
import exceltools.com.excel.base.ExcelUtils;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

/**
 * @author sunc
 * @date 2019/11/13 19:14
 * @description MergeHandler
 */

public class MergeHandler {

    public static String merge(File file) {
        try {
            return mergeExcel(file);
        } catch (OpenXML4JException | ParserConfigurationException | SAXException | IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    private static String mergeExcel(File file) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        String destination = null;
        if (file.exists()) {
            if (file.isDirectory()) {
                destination = file.getAbsolutePath() + File.separator + System.currentTimeMillis() + "_merge.xls";
                File[] list = file.listFiles((dir, name) -> Common.validExcelType(name));
                if (list != null) {
                    List<File> files = Arrays.asList(list);
                    ExcelUtils.writeExcel(destination, files);
                }
            }
        }
        return destination;
    }

}
