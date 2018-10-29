package indi.liht.stat.utils;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

/**
 * Usage:
 * 日期工具类
 * @author lihongtao ibraxwell@sina.com
 * on 2018/10/29
 **/
public class DateUtils {

    /** 纯数字的年月日时分秒的日期格式 */
    public static final String yyyyMMddHHmmss = "yyyyMMddHHmmss";

    /**
     * 日期字符串
     * @param date 日期
     * @param pattern 日期格式
     * @return 对应的日期字符串
     */
    public static String getStrFromDate(Date date, String pattern) {
        DateFormat df = new SimpleDateFormat(pattern, Locale.US);
        return df.format(date);
    }

}
