package exceltools.com.excel.base;

import exceltools.com.excel.vo.CellContainer;
import org.apache.commons.lang3.time.FastDateFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

/**
 * @author sunc
 * @date 2019/11/13 18:53
 * @description ExcelUtils
 */

public class ExcelUtils {
    /**
     * 格式化科学计数器
     */
    private static final DecimalFormat DECIMAL_FORMAT_NUMBER = new DecimalFormat("0.00E000");
    /**
     * 小数匹配
     */
    private static final Pattern POINTS_PATTERN = Pattern.compile("0.0+_*[^/s]+");

    public static void writeExcelString(String destination, List<File> files) throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {
        Workbook workbook = createWorkbookIfNotExist(destination);
        // 获取文件的指定工作表 默认的第一个
        Sheet sheet = workbook.createSheet();

        List<String[]> lines;
        int rowNumber = 0;
        for (File file : files) {
            lines = readExcelString(file);
            for (String[] line : lines) {
                Row row = sheet.createRow(rowNumber++);
                int cellNumber = 0;
                for (String val : line) {
                    Cell cell = row.createCell(cellNumber++);
                    cell.setCellValue(val);
                }
            }

        }
        FileOutputStream fos = new FileOutputStream(destination);
        workbook.write(fos);
        fos.close();
        workbook.close();
    }

    public static void writeExcelCell(String destination, List<File> files) throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {
        Workbook workbook = createWorkbookIfNotExist(destination);
        // 获取文件的指定工作表 默认的第一个
        Sheet sheet = workbook.createSheet();

        List<List<CellContainer>> lines = new ArrayList<>();
        int rowNumber = 0;

        List<CellContainer> header = new ArrayList<>();
        for (File file : files) {
            List<List<CellContainer>> tmp = readExcelCell(file);
            List<CellContainer> headerTmp = tmp.remove(0);
            if (header.size() < headerTmp.size()) {
                header = headerTmp;
            }
            lines.addAll(tmp);
        }
        if (!header.isEmpty()) {
            lines.add(0, header);
        }
        for (List<CellContainer> line : lines) {
            Row row = sheet.createRow(rowNumber++);
            int cellNumber = 0;
            for (CellContainer val : line) {
                Cell cell = row.createCell(cellNumber++);
                if (val == null) {
                    cell.setCellType(CellType.STRING);
                    cell.setCellValue("");
                } else {
                    val.setCell(cell);
                }
            }
        }
        FileOutputStream fos = new FileOutputStream(destination);
        workbook.write(fos);
        fos.flush();
        fos.close();
        workbook.close();
    }

    public static void splitExcelCell(String destination, File file) throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {

    }

    private static Workbook createWorkbookIfNotExist(String fileName) throws IOException {
        Workbook wb = new HSSFWorkbook();

        try (OutputStream output = new FileOutputStream(fileName)) {
            wb.write(output);
        } catch (FileNotFoundException e) {
            throw new FileNotFoundException();
        }

        return wb;
    }

    private static List<String[]> readExcelString(File file) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        String type = file.getAbsolutePath();
        if (type.endsWith("xls")) {
            return readXlsString(file);
        }
        return readXlsxString(file);
    }

    private static List<List<CellContainer>> readExcelCell(File file) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        String type = file.getAbsolutePath();
        if (type.endsWith("xls")) {
            return readXlsCell(file);
        }
        return readXlsxCell(file);
    }

    /**
     * 读取 xls
     *
     * @param file file
     * @return xls 内容
     */
    private static List<String[]> readXlsString(File file) throws IOException, InvalidFormatException {
        List<String[]> res = new ArrayList<>();

        InputStream inputStream = new BufferedInputStream(new FileInputStream(file));
        Workbook workbook = WorkbookFactory.create(inputStream);
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        // 获取文件的指定工作表 默认的第一个
        Sheet sheet = workbook.getSheetAt(0);
        int rowCount = sheet.getPhysicalNumberOfRows();
        int index = 0;
        Row head = sheet.getRow(index++);
        while (head == null && index < rowCount) {
            head = sheet.getRow(index++);
        }
        if (head == null) {
            return res;
        }
        int length = head.getPhysicalNumberOfCells();
        String[] line = Common.init(length);
        for (int j = 0; j < length; j++) {
            Cell cell = head.getCell(j);
            if (cell != null) {
                line[j] = getCellValue(evaluator, cell);
            }
        }
        res.add(line);

        for (int i = index; i < rowCount; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            line = Common.init(length);
            for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                Cell cell = row.getCell(j);
                if (cell != null) {
                    line[j] = getCellValue(evaluator, cell);
                }
            }
            res.add(line);
        }
        workbook.close();
        inputStream.close();
        return res;
    }

    /**
     * 读取 xlsx
     *
     * @param file file
     * @return xlsx 内容
     */
    private static List<String[]> readXlsxString(File file) throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {
        List<String[]> res = new ArrayList<>();

        OPCPackage p = OPCPackage.open(file, PackageAccess.READ);
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(p);

        XSSFReader xssfReader = new XSSFReader(p);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        InputStream stream = iter.next();

        InputSource sheetSource = new InputSource(stream);

        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
        SAXParser saxParser = saxFactory.newSAXParser();
        XMLReader parser = saxParser.getXMLReader();

        MyDefaultHandler handler = new MyDefaultHandler(strings, styles);
        parser.setContentHandler(handler);
        parser.parse(sheetSource);

        List<List<String>> rows = handler.getRows();
        if (rows.isEmpty()) {
            return res;
        }
        int length = rows.get(0).size();
        for (List<String> row : rows) {
            int loop = Math.min(row.size(), length);
            String[] line = Common.init(length);
            for (int j = 0; j < loop; j++) {
                line[j] = row.get(j);
            }
            res.add(line);
        }
        stream.close();
        p.close();
        return res;
    }

    /**
     * 读取 xls
     *
     * @param file file
     * @return xls 内容
     */
    private static List<List<CellContainer>> readXlsCell(File file) throws IOException, InvalidFormatException {
        List<List<CellContainer>> res = new ArrayList<>();

        InputStream inputStream = new BufferedInputStream(new FileInputStream(file));
        Workbook workbook = WorkbookFactory.create(inputStream);
        // 获取文件的指定工作表 默认的第一个
        Sheet sheet = workbook.getSheetAt(0);
        int rowCount = sheet.getPhysicalNumberOfRows();
        int index = 0;
        Row head = sheet.getRow(index++);
        while (head == null && index < rowCount) {
            head = sheet.getRow(index++);
        }
        if (head == null) {
            return res;
        }
        int length = head.getPhysicalNumberOfCells();
        List<CellContainer> lineHeader = new ArrayList<>();
        for (int j = 0; j < length; j++) {
            Cell cell = head.getCell(j);
            lineHeader.add(getCell(cell));
        }
        res.add(lineHeader);

        for (int i = index; i < rowCount; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            List<CellContainer> line = new ArrayList<>();
            for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                Cell cell = row.getCell(j);
                line.add(getCell(cell));
            }
            res.add(line);
        }
        workbook.close();
        inputStream.close();
        return res;
    }

    /**
     * 读取 xlsx
     *
     * @param file file
     * @return xlsx 内容
     */
    private static List<List<CellContainer>> readXlsxCell(File file) throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {
        OPCPackage p = OPCPackage.open(file, PackageAccess.READ);
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(p);

        XSSFReader xssfReader = new XSSFReader(p);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        InputStream stream = iter.next();

        InputSource sheetSource = new InputSource(stream);

        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
        SAXParser saxParser = saxFactory.newSAXParser();
        XMLReader parser = saxParser.getXMLReader();

        MyCellHandler handler = new MyCellHandler(strings, styles);
        parser.setContentHandler(handler);
        parser.parse(sheetSource);
        List<List<CellContainer>> rows = handler.getRows();
        stream.close();
        p.close();
        return rows;
    }

    private static String getCellValue(FormulaEvaluator evaluator, Cell cell) {
        Object value = "";
        if (cell == null) {
            return "";
        }
        CellType type = cell.getCellTypeEnum();
        switch (type) {
            case FORMULA:
                CellValue cellValue = evaluator.evaluate(cell);
                value = formulaCellValue(cellValue);
                break;
            case _NONE:
                break;
            case STRING:
                value = cell.getStringCellValue();
                break;
            case NUMERIC:
                String dataFormatString = cell.getCellStyle().getDataFormatString();
                if (DateUtil.isCellDateFormatted(cell)) {
                    //日期
                    value = FastDateFormat.getInstance(Constant.DATE_FORMAT).format(DateUtil.getJavaDate(cell.getNumericCellValue()));
                } else if ("@".equals(dataFormatString)
                        || "General".equals(dataFormatString)
                        || "0_ ".equals(dataFormatString)) {
                    //文本  or 常规 or 整型数值
                    Double tmp = cell.getNumericCellValue();
                    if (Double.parseDouble(tmp.intValue() + ".0") == tmp) {
                        value = tmp.intValue();
                    } else {
                        value = tmp;
                    }
                } else if ("0.00%".equals(dataFormatString)) {
                    //百分比
                    value = cell.getNumericCellValue();
                    value = new DecimalFormat("##.00%").format(value);
                } else if ("0%".equals(dataFormatString)) {
                    //百分比
                    value = cell.getNumericCellValue();
                    value = new DecimalFormat("##%").format(value);
                } else if (POINTS_PATTERN.matcher(dataFormatString).matches()) {
                    //正则匹配小数类型
                    value = cell.getNumericCellValue();
                } else if ("0.00E+00".equals(dataFormatString)) {
                    //科学计数
                    value = cell.getNumericCellValue();
                    value = DECIMAL_FORMAT_NUMBER.format(value);
                } else if ("# ?/?".equals(dataFormatString)) {
                    //分数
                    value = cell.getNumericCellValue();
                } else {
                    //货币
                    value = cell.getNumericCellValue();
                }
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case BLANK:
                //value = "";
                break;
            default:
                value = cell.toString();
        }
        return String.valueOf(value).replaceAll("(\r\n|\r|\n|\n\r)", " ");
    }

    private static Object formulaCellValue(CellValue cellValue) {
        Object value = "";
        switch (cellValue.getCellTypeEnum()) {
            case NUMERIC:
                value = cellValue.getNumberValue();
                break;
            case STRING:
                value = cellValue.getStringValue();
                break;
            case BOOLEAN:
                value = cellValue.getBooleanValue();
                break;
            case ERROR:
                value = cellValue.getErrorValue();
                break;
            default:
        }
        return value;
    }

    private static CellContainer getCell(Cell cell) {
        Object value = "";
        String type = CellContainer.NONE;
        if (cell == null) {
            return new CellContainer(null, null);
        }
        CellType cellType = cell.getCellTypeEnum();
        switch (cellType) {
            case FORMULA:
                type = CellContainer.FORMULA;
                value = cell.getCellFormula();
                break;
            case _NONE:
                break;
            case STRING:
                type = CellContainer.STRING;
                value = cell.getStringCellValue();
                break;
            case NUMERIC:
                type = CellContainer.NUMBER;
                String dataFormatString = cell.getCellStyle().getDataFormatString();
                if (DateUtil.isCellDateFormatted(cell)) {
                    //日期
                    type = CellContainer.DATE;
                    value = FastDateFormat.getInstance(Constant.DATE_FORMAT).format(DateUtil.getJavaDate(cell.getNumericCellValue()));
                } else if ("@".equals(dataFormatString)
                        || "General".equals(dataFormatString)
                        || "0_ ".equals(dataFormatString)) {
                    //文本  or 常规 or 整型数值
                    Double tmp = cell.getNumericCellValue();
                    if (Double.parseDouble(tmp.intValue() + ".0") == tmp) {
                        value = tmp.intValue();
                    } else {
                        value = tmp;
                    }
                } else if ("0.00%".equals(dataFormatString)) {
                    //百分比
                    value = cell.getNumericCellValue();
                    value = new DecimalFormat("##.00%").format(value);
                } else if ("0%".equals(dataFormatString)) {
                    //百分比
                    value = cell.getNumericCellValue();
                    value = new DecimalFormat("##%").format(value);
                } else if (POINTS_PATTERN.matcher(dataFormatString).matches()) {
                    //正则匹配小数类型
                    value = cell.getNumericCellValue();
                } else if ("0.00E+00".equals(dataFormatString)) {
                    //科学计数
                    value = cell.getNumericCellValue();
                    value = DECIMAL_FORMAT_NUMBER.format(value);
                } else if ("# ?/?".equals(dataFormatString)) {
                    //分数
                    value = cell.getNumericCellValue();
                } else {
                    //货币
                    value = cell.getNumericCellValue();
                }
                break;
            case BOOLEAN:
                type = CellContainer.BOOLEAN;
                value = cell.getBooleanCellValue();
                break;
            case BLANK:
                //value = "";
                break;
            default:
                value = cell.toString();
        }
        return new CellContainer(type, String.valueOf(value));
    }


}
