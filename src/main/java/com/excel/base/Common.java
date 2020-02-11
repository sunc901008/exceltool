package com.excel.base;

import java.io.File;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author sunc
 * @date 2019/11/13 18:49
 * @description Common
 */

public class Common {

    public static boolean validExcelType(File file) {
        if (file.isDirectory()) {
            return true;
        }
        return validExcelType(file.getAbsoluteFile());
    }

    public static boolean validExcelType(String filePath) {
        String type = filePath.substring(filePath.lastIndexOf(".") + 1).toLowerCase();
        return Constant.TYPES.contains(type);
    }

    public static boolean isEmpty(String str) {
        return str == null || "".equals(str.trim());
    }

    static String[] init(int length) {
        String[] ss = new String[length];
        for (int i = 0; i < length; i++) {
            ss[i] = null;
        }
        return ss;
    }

    static String datetimeFormat(Date date) {
        SimpleDateFormat sdf = new SimpleDateFormat(Constant.DATE_FORMAT);
        return sdf.format(date);
    }

    public static Date datetimeParse(String date) {
        SimpleDateFormat sdf = new SimpleDateFormat(Constant.DATE_FORMAT);
        try {
            return sdf.parse(date);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return new Date();
    }

}
