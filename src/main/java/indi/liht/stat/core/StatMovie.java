package indi.liht.stat.core;

import indi.liht.stat.constants.StatConsts;
import indi.liht.stat.models.MovieComSection;
import indi.liht.stat.utils.DateUtils;
import indi.liht.stat.utils.PoiUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.concurrent.*;
import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReentrantLock;

/**
 * Usage:
 * 统计电影情况Excel核心类
 * @author lihongtao ibraxwell@sina.com
 * on 2018/10/28
 **/
public class StatMovie implements IStat {

    /** 资源文件路径 */
    private static String RESOURCE_ROOT_PATH =
            StatMovie.class.getProtectionDomain().getCodeSource().getLocation().getPath();

    /** Excel文件路径 */
    private static String EXCEL_ROOT_PATH;

    /** stat.properties配置文件 */
    private Properties properties = new Properties();

    static {
        // 打包成jar后，读取外层同级的配置文件和Excel文件
        if (RESOURCE_ROOT_PATH.toLowerCase().endsWith(".jar")) {
            int lastIndex = RESOURCE_ROOT_PATH.lastIndexOf("/");
            RESOURCE_ROOT_PATH = RESOURCE_ROOT_PATH.substring(0, lastIndex) + "/";
        }
        EXCEL_ROOT_PATH = RESOURCE_ROOT_PATH + "conf/excel/";
    }

    /** 要写入的工作表 */
    private Workbook workbookToWrite;

    /** 用于createSheet时的锁 */
    private Lock lock = new ReentrantLock();

    /**
     * 构造函数
     */
    public StatMovie() {}

    /**
     * 主方法
     */
    public void stat() {
        this.loadProps();
        this.handleData();
    }

    /**
     * 加载stat.properties配置文件
     */
    private void loadProps() {
        InputStream is = null;
        Reader isr = null;
        String propsFilePath = RESOURCE_ROOT_PATH + "conf/stat.properties";
        try {
            is = new FileInputStream(propsFilePath);
            isr = new InputStreamReader(is, StandardCharsets.UTF_8);
            properties.load(isr);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(is);
            IOUtils.closeQuietly(isr);
        }
        System.out.println("加载[" + propsFilePath + "]配置文件完成！");
    }

    /**
     * 处理赋值Excel数据，然后进入下一环节
     */
    private void handleData() {
        Sheet[] relationDataSheet = PoiUtils.getSheetsFromPath(EXCEL_ROOT_PATH
                        + properties.getProperty("stat.movie.relation-data.file-name").trim(),
                properties.getProperty("stat.movie.relation-data.sheet-name").trim());
        if (Optional.ofNullable(relationDataSheet).isPresent()
                && Optional.ofNullable(relationDataSheet[0]).isPresent()) {
            // A+B->weight
            Map<String, Integer> sectionABWeight = new HashMap<>();
            int cnt = 0;
            int rowCount = relationDataSheet[0].getLastRowNum() - relationDataSheet[0].getFirstRowNum() + 1;
            for (int rowNum = 1; rowNum < rowCount; rowNum++) {
                Row row = relationDataSheet[0].getRow(rowNum);
                String sectionA = String.valueOf(PoiUtils.getCellValueFromRow(row, 0));
                String sectionB = String.valueOf(PoiUtils.getCellValueFromRow(row, 1));
                int weight = Double.valueOf(Optional.ofNullable(
                        String.valueOf(PoiUtils.getCellValueFromRow(row, 2))).orElse("0"))
                        .intValue();
                if (StringUtils.isNotEmpty(sectionA) && StringUtils.isNotEmpty(sectionB)) {
                    String key = sectionA + StatConsts.KEY_SEPARATOR + sectionB;
                    sectionABWeight.put(key, weight);
                    cnt++;
                }
            }
            System.out.println("共加载[" + cnt + "]行有效数据！");
            // 处理主数据来源Excel数据
            this.handleSourceData(sectionABWeight);
        } else {
            System.out.println("加载赋值Excel数据失败！请检查stat.movie.relation-data的相关配置项！");
        }
    }

    /**
     * 处理来源Excel数据
     * @param sectionABWeight 赋值的分工权重
     */
    private void handleSourceData(Map<String, Integer> sectionABWeight) {
        // 多个Sheet名字，split分割后，用数组存起来
        String[] sourceDataSheetNames =
                properties.getProperty("stat.movie.main-data.sheet-names").trim()
                        .split(StatConsts.PROPS_VALUE_SEPARATOR);
        // 多个Sheet，也是用数组存起来
        Sheet[] sourceDataSheet = PoiUtils.getSheetsFromPath(
                EXCEL_ROOT_PATH + properties.getProperty("stat.movie.main-data.file-name").trim(),
                sourceDataSheetNames);

        // 生成文件名，然后根据Excel文件类型构造对应的Workbook
        String outputFileName = "stat-movie-"
                + DateUtils.getStrFromDate(new Date(), DateUtils.yyyyMMddHHmmss);
        String outputFileSuffix = properties.getProperty("stat.movie.output-data.suffix")
                .trim().toLowerCase();
        switch (outputFileSuffix) {
            case "xls":
                workbookToWrite = new HSSFWorkbook();
                outputFileName += ".xls";
                break;
            default:
                workbookToWrite = new XSSFWorkbook();
                outputFileName += ".xlsx";
        }

        // 用于存放汇总的电影公司权重
        Map<String, Integer> allMovieABWeight = new ConcurrentHashMap<>();
        // Sheet的数量
        int sheetSize = sourceDataSheetNames.length;
        // 固定大小线程池
        ExecutorService threadPool = Executors.newFixedThreadPool(sheetSize);
        // CountDownLatch
        final CountDownLatch latch = new CountDownLatch(sheetSize);
        for (int i = 0; i < sheetSize; i++) {
            if (Optional.ofNullable(sourceDataSheet).isPresent()
                    && Optional.ofNullable(sourceDataSheet[i]).isPresent()) {

                // 多线程，一个Sheet一个线程去解析，并且生成结果Sheet。结果Sheet名跟源Sheet名一样
                final int _i = i;
                threadPool.submit(() -> {
                    Thread.currentThread().setName(sourceDataSheetNames[_i] +"-Thread");
                    // A||B->weight
                    Map<String, Integer> movieABWeight = new HashMap<>();
                    int cnt = 0;
                    int rowCount = sourceDataSheet[_i].getLastRowNum()
                            - sourceDataSheet[_i].getFirstRowNum();
                    for (int rowNum = 0; rowNum < rowCount + 1; rowNum++) {
                        // 存放电影公司与其分工
                        List<MovieComSection> movieComSectionList = new LinkedList<>();
                        if (isUselessRow(rowNum)) {
                            continue;
                        }
                        boolean useful = false;
                        Row row = sourceDataSheet[_i].getRow(rowNum);
                        int cellNum = row.getLastCellNum() - row.getFirstCellNum();
                        for (int columnNum = 0; columnNum < cellNum; columnNum++) {
                            String section = String.valueOf(PoiUtils.getCellValueFromRow(row, columnNum));
                            String movieCom = String.valueOf(
                                    PoiUtils.getCellValueFromRow(sourceDataSheet[_i].getRow(rowNum + 1), columnNum));
                            if (StringUtils.isNotEmpty(section) && StringUtils.isNotEmpty(movieCom)) {
                                movieComSectionList.add(new MovieComSection(movieCom, section));
                                if (!useful) {
                                    useful = true;
                                }
                            }
                        }
                        if (useful) {
                            cnt++;
                        }
                        // 根据movieComSectionList和sectionABWeight计算权重
                        putToWeightMap(movieComSectionList, movieABWeight, sectionABWeight);
                    }
                    // 合并单个权重到汇总权重
                    putPartToAll(movieABWeight, allMovieABWeight);
                    System.out.println("共加载[" + cnt * 2 + "]行有效数据！");
                    // 用计算后的结果，生成结果Sheet
                    handleOutputData(movieABWeight, sourceDataSheetNames[_i], outputFileSuffix);
                    // CountDownLatch倒数
                    latch.countDown();
                    System.out.println("[" + Thread.currentThread().getName() + "]线程执行完成！");
                });

            } else {
                System.out.println("加载来源Excel数据失败！请检查stat.movie.main-data的相关配置项");
            }
        }
        // 等待所有Sheet完成，才能生成汇总的Sheet
        try {
            System.out.println("等待" + sheetSize + "个子线程执行完毕……");
            latch.await();
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        // shutdown线程池
        threadPool.shutdown();

        // 汇总
        this.handleOutputData(allMovieABWeight, "All", outputFileSuffix);

        // 保存Workbook工作表
        OutputStream os = null;
        try {
            os = new FileOutputStream(RESOURCE_ROOT_PATH + outputFileName);
            workbookToWrite.write(os);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(os);
        }
    }

    /**
     * 把单个Sheet的电影公司权重，合并到总的权重中去
     * @param movieABWeight 个Sheet的电影公司权重
     * @param allMovieABWeight 总的权重
     */
    private void putPartToAll(Map<String, Integer> movieABWeight,
                              Map<String, Integer> allMovieABWeight) {
        for (Map.Entry<String, Integer> entry : movieABWeight.entrySet()) {
            String originalKey = entry.getKey();
            int weight = entry.getValue();
            if (allMovieABWeight.containsKey(originalKey)) {
                allMovieABWeight.put(originalKey, weight + allMovieABWeight.get(originalKey));
            } else {
                String[] keyPart = originalKey.split(StatConsts.KEY_SEPARATOR);
                String newKey = keyPart[1] + StatConsts.KEY_SEPARATOR + keyPart[0];
                if (allMovieABWeight.containsKey(newKey)) {
                    allMovieABWeight.put(newKey, weight + allMovieABWeight.get(newKey));
                } else {
                    allMovieABWeight.put(newKey, weight);
                }
            }
        }
    }

    /**
     * 计算电影公司权重
     * @param movieComSectionList 电影公司与其分工
     * @param movieABWeight 返回的电影公司权重
     * @param sectionABWeight 分工权重
     */
    private void putToWeightMap(List<MovieComSection> movieComSectionList,
                                Map<String, Integer> movieABWeight,
                                Map<String, Integer> sectionABWeight) {
        for (int i = 0, listSize = movieComSectionList.size(); i < listSize; i++) {
            for (int j = i + 1; j < listSize; j++) {
                // 根据i和j两个电影公司的分工，确定分工权重
                int weight;
                String weightKey = movieComSectionList.get(i).getSection()
                        + StatConsts.KEY_SEPARATOR
                        + movieComSectionList.get(j).getSection();
                if (!sectionABWeight.containsKey(weightKey)) {
                    weightKey = movieComSectionList.get(j).getSection()
                            + StatConsts.KEY_SEPARATOR
                            + movieComSectionList.get(i).getSection();
                    weight = sectionABWeight.getOrDefault(weightKey, 0);
                } else {
                    weight = sectionABWeight.get(weightKey);
                }
                // 加上上面计算的分工权重。需要判断是否已经存在，存在则累加权重
                String originalMovieKey = movieComSectionList.get(i).getMovieCom()
                        + StatConsts.KEY_SEPARATOR
                        + movieComSectionList.get(j).getMovieCom();
                if (!movieABWeight.containsKey(originalMovieKey)) {
                    String newMovieKey = movieComSectionList.get(j).getMovieCom()
                            + StatConsts.KEY_SEPARATOR
                            + movieComSectionList.get(i).getMovieCom();
                    if (!movieABWeight.containsKey(newMovieKey)) {
                        movieABWeight.put(newMovieKey, weight);
                    } else {
                        movieABWeight.put(newMovieKey, movieABWeight.get(newMovieKey) + weight);
                    }
                } else {
                    movieABWeight.put(originalMovieKey, movieABWeight.get(originalMovieKey) + weight);
                }
            }
        }
    }

    /**
     * 根据行号去读源数据Excel
     * @param rowNum 行号
     * @return 是否需要读
     */
    private boolean isUselessRow(int rowNum) {
        // 只取3的倍数行
        return (rowNum + 1) % 4 != 3;
    }

    /**
     * 输出结果到新的Sheet
     * @param movieABWeight 电影公司权重
     * @param sheetName 输出的Sheet名
     * @param outputFileSuffix 输出Excel后缀
     */
    private void handleOutputData(Map<String, Integer> movieABWeight,
                                  String sheetName, String outputFileSuffix) {
        // 用于辅助movieComList去重
        Set<String> movieComSet = new HashSet<>();
        // 用于生成第一行和第一列
        List<String> movieComList = new LinkedList<>();
        for (String key : movieABWeight.keySet()) {
            String[] keyParts = key.split(StatConsts.KEY_SEPARATOR);
            String movieComA = keyParts[0];
            String movieComB = keyParts[1];
            if (!movieComSet.contains(movieComA)) {
                movieComSet.add(movieComA);
                movieComList.add(movieComA);
            }
            if (!movieComSet.contains(movieComB)) {
                movieComSet.add(movieComB);
                movieComList.add(movieComB);
            }
        }

        // 根据后缀区分要Excel写入类型
        switch (outputFileSuffix) {
            case "xls":
                this.writeToHSSFSheet(sheetName, movieABWeight, movieComList);
                break;
            default:
                this.writeToXSSFSheet(sheetName, movieABWeight, movieComList);
        }

    }

    /**
     * 输出到xls后缀的Excel中
     * @param sheetName 输出的Sheet名
     * @param movieABWeight 电影公司权重
     * @param movieComList 涉及到的所有电影公司列表
     */
    private void writeToHSSFSheet(String sheetName,
                              Map<String, Integer> movieABWeight,
                              List<String> movieComList) {
        HSSFSheet sheet;
        lock.lock();
        try {
            sheet = (HSSFSheet) workbookToWrite.createSheet(sheetName);
        } finally {
            lock.unlock();
        }

        // 写入第一行
        HSSFRow firstRow = sheet.createRow(0);
        // 第一行第一列空白
        firstRow.createCell(0).setCellValue("");
        for (int i = 0, listSize = movieComList.size(); i < listSize; i++) {
            firstRow.createCell(i + 1).setCellValue(movieComList.get(i));
        }

        // 写入第二行开始的行
        for (int i = 0, listSize = movieComList.size(); i < listSize; i++) {
            // 每行第一列写入公司名
            HSSFRow row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(movieComList.get(i));
            this.writeWeightToCell(i, listSize, row, movieABWeight, movieComList);
        }
        System.out.println("完成写入[" + sheetName + "]工作簿[类型：xls]，共"
                + (movieComList.size() + 1) + "行！");
    }

    /**
     * 输出到xlsx后缀的Excel中
     * @param sheetName 输出的Sheet名
     * @param movieABWeight 电影公司权重
     * @param movieComList 涉及到的所有电影公司列表
     */
    private void writeToXSSFSheet(String sheetName,
                              Map<String, Integer> movieABWeight,
                              List<String> movieComList) {
        XSSFSheet sheet;
        lock.lock();
        try {
            sheet = (XSSFSheet) workbookToWrite.createSheet(sheetName);
        } finally {
            lock.unlock();
        }

        // 写入第一行
        XSSFRow firstRow = sheet.createRow(0);
        // 第一行第一列空白
        firstRow.createCell(0).setCellValue("");
        for (int i = 0, listSize = movieComList.size(); i < listSize; i++) {
            firstRow.createCell(i + 1).setCellValue(movieComList.get(i));
        }

        // 写入第二行开始的行
        for (int i = 0, listSize = movieComList.size(); i < listSize; i++) {
            // 每行第一列写入公司名
            XSSFRow row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(movieComList.get(i));
            this.writeWeightToCell(i, listSize, row, movieABWeight, movieComList);
        }
        System.out.println("完成写入[" + sheetName + "]工作簿[类型：xlsx]，共"
                + (movieComList.size() + 1) + "行！");
    }

    /**
     * 把权重写入
     * @param i 行数
     * @param listSize 公司列表数量
     * @param row 当前行
     * @param movieABWeight 电影公司权重
     * @param movieComList 涉及到的所有电影公司列表
     */
    private void writeWeightToCell(int i, int listSize, Row row,
                                   Map<String, Integer> movieABWeight,
                                   List<String> movieComList) {
        for (int j = 0; j < listSize; j++) {
            if (i == j) {
                row.createCell(j + 1).setCellValue(0);
            }
            String key = movieComList.get(i)
                    + StatConsts.KEY_SEPARATOR
                    + movieComList.get(j);
            if (movieABWeight.containsKey(key)) {
                row.createCell(j + 1).setCellValue(movieABWeight.get(key));
            } else {
                // 把分隔符两边对调作为新的Key再去Map查
                key = movieComList.get(j)
                        + StatConsts.KEY_SEPARATOR
                        + movieComList.get(i);
                row.createCell(j + 1).setCellValue(movieABWeight.getOrDefault(key, 0));
            }
        }
    }

}
