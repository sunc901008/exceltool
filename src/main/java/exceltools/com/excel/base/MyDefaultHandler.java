package exceltools.com.excel.base;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static exceltools.com.excel.base.MyDefaultHandler.xssfDataType.SSTINDEX;

/**
 * @author sunc
 * @date 2019/11/13 19:04
 * @description MyDefaultHandler
 */

public class MyDefaultHandler extends DefaultHandler {
    private static final Logger logger = Logger.getLogger(MyDefaultHandler.class);

    private boolean vIsOpen;
    private StringBuffer value = new StringBuffer();
    private xssfDataType nextDataType;
    private ReadOnlySharedStringsTable strings;
    private StylesTable styles;
    private short formatIndex;
    private String formatString;
    private DataFormatter formatter = new DataFormatter();
    private List<String> record = new ArrayList<>();
    private List<List<String>> rows = new ArrayList<>();
    private int rowNum = 1;
    private int colNum = 0;

    // 定义当前读到的列数，实际读取时会按照从0开始...
    private int thisColumn = -1;
    // 定义上一次读到的列序号
    private int lastColumnNumber = -1;

    enum xssfDataType {
        BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER,
    }

    public MyDefaultHandler(ReadOnlySharedStringsTable strings, StylesTable styles) {
        this.strings = strings;
        this.styles = styles;
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        if ("inlineStr".equals(qName) || "v".equals(qName)) {
            vIsOpen = true;
            // Clear contents cache
            value.setLength(0);
        } else if ("c".equals(qName)) {

            // Get the cell reference
            String r = attributes.getValue("r");
            int firstDigit = -1;
            for (int c = 0; c < r.length(); ++c) {
                if (Character.isDigit(r.charAt(c))) {
                    firstDigit = c;
                    break;
                }
            }
            thisColumn = nameToColumn(r.substring(0, firstDigit));//获取当前读取的列数

            nextDataType = xssfDataType.NUMBER;
            formatIndex = -1;
            formatString = null;
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");
            if ("b".equals(cellType)) {
                nextDataType = xssfDataType.BOOL;
            } else if ("e".equals(cellType)) {
                nextDataType = xssfDataType.ERROR;
            } else if ("inlineStr".equals(cellType)) {
                nextDataType = xssfDataType.INLINESTR;
            } else if ("s".equals(cellType)) {
                nextDataType = SSTINDEX;
            } else if ("str".equals(cellType)) {
                nextDataType = xssfDataType.FORMULA;
            } else if (cellStyleStr != null) {
                int styleIndex = Integer.parseInt(cellStyleStr);
                XSSFCellStyle style = styles.getStyleAt(styleIndex);
                this.formatIndex = style.getDataFormat();
                this.formatString = style.getDataFormatString();
                if (this.formatString == null) {
                    this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
                }
            }
        }

    }

    @Override
    public void characters(char[] ch, int start, int length) {
        if (vIsOpen) {
            value.append(ch, start, length);
        }
    }

    private int nameToColumn(String name) {
        int column = -1;
        for (int i = 0; i < name.length(); ++i) {
            int c = name.charAt(i);
            column = (column + 1) * 26 + c - 'A';
        }
        return column;
    }

    @Override
    public void endElement(String uri, String localName, String name) {
        if ("v".equals(name)) {
            colNum++;
            String thisStr = "";
            switch (nextDataType) {
                case BOOL:
                    char first = value.charAt(0);
                    thisStr = first == '0' ? "FALSE" : "TRUE";
                    break;
                case ERROR:
                    logger.error(String.format("ERROR:%s ROW NUMBER:%s COL NUMBER:%s", value.toString(), rowNum, colNum));
                    break;
                case FORMULA:
                    logger.error(String.format("This is a formula string::%s ROW NUMBER:%s COL NUMBER:%s", value.toString(), rowNum, colNum));
                    break;
                case INLINESTR:
                    XSSFRichTextString ts = new XSSFRichTextString(value.toString());
                    thisStr = ts.toString();
                    break;
                case SSTINDEX:
                    String sstIndex = value.toString();
                    try {
                        int idx = Integer.parseInt(sstIndex);
                        XSSFRichTextString rtss = new XSSFRichTextString(strings.getEntryAt(idx));
                        thisStr = rtss.toString();
                    } catch (NumberFormatException ex) {
                        logger.error(String.format("Failed to parse SST index '" + sstIndex + "': " + ex.toString() + "ROW NUMBER:%s COL NUMBER:%s", rowNum, colNum));
                    }
                    break;
                case NUMBER:
                    String n = value.toString();
                    // 判断是否是日期格式
                    if (HSSFDateUtil.isADateFormat(this.formatIndex, n)) {
                        Double d = Double.parseDouble(n);
                        Date date = HSSFDateUtil.getJavaDate(d);
                        thisStr = Common.datetimeFormat(date);
                    } else if (this.formatString != null) {
                        thisStr = formatter.formatRawCellContents(Double.parseDouble(n), this.formatIndex, this.formatString);
                    } else {
                        thisStr = n;
                    }
                    break;

                default:
                    logger.error(String.format("Unexpected type:%s VALUE:%s ROW NUMBER:%s COL NUMBER:%s", nextDataType, value.toString(), rowNum, colNum));
                    thisStr = value.toString();
                    break;
            }
            int addSpace = thisColumn - lastColumnNumber;
            while (addSpace > 1) {
                record.add(null);
                addSpace--;
            }
            thisStr = thisStr.replaceAll("(\r\n|\r|\n|\n\r)", " ");
            record.add(thisStr);
            if (thisColumn > -1) {
                lastColumnNumber = thisColumn;
            }
        } else if ("row".equals(name)) {
            colNum = 0;
            lastColumnNumber = -1;
            if (record.isEmpty()) {
                return;
            }
            rows.add(new ArrayList<>(record));
            rowNum++;
            record.clear();
        }

    }

    public List<List<String>> getRows() {
        return rows;
    }

}
