package com.excel;

import com.excel.base.ExcelUtils;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.util.Arrays;

/**
 * @author sunc
 * @date 2020/2/11 17:45
 */
public class Home {

    public static void main(String[] args) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        File fileDir = new File("G:\\Temp");
        String destination = "G:\\Temp\\merge.xls";
        if (!fileDir.exists()) {
            return;
        }
        File[] files = fileDir.listFiles();
        if (files != null) {
            ExcelUtils.writeExcelCell(destination, Arrays.asList(files));
        }

    }

}
