package com.Noah.utils;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.omg.CORBA.UserException;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Slf4j
public class ExcelTools {

    //从第几行开始读取，默认从第一行读取
    private static int readStartRowPos = 1;

    //数据处理逻辑
    @FunctionalInterface
    public interface LogicCalculator {
        LogicResult calculate(List<Map<String, Object>> resultList, String resultKey);
    }

    @Data
    public static class LogicResult {

        private String fileAuthor = "system";

        private String fileExt = "xls";

        private String fileName = "导入报告";

        private boolean result = true;

        private List<Map<String, Object>> list = new ArrayList<>();

        public LogicResult() {
            super();
        }

        public LogicResult(boolean result, List<Map<String, Object>> list) {
            this.result = result;
            this.list = list;
        }
    }




    /**
     * 写入本地Excel文件
     *
     * @param file 文件
     * @param wb   数据
     */
    public static void writeLocalFile(File file, HSSFWorkbook wb) {
        //判断建立文件
        if (!file.exists()) {
            try {
                file.createNewFile();
            } catch (IOException e) {
                //创建文件失败
                log.error("创建文件失败", e);
                return;
            }
        }
        //使用true，即进行append file
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(file, false);
            wb.write(fileOutputStream);
            fileOutputStream.flush();
        } catch (Exception e) {
            log.error("文件写入异常", e);
        } finally {
            if (fileOutputStream != null) {
                try {
                    fileOutputStream.close();
                } catch (IOException e) {
                    //nothing to do
                }
            }
        }
    }

    /**
     * 获取文件的字节数组
     *
     * @param hssfWorkbook 表信息
     * @return 字节数组
     */
    public static byte[] getExcelBytes(HSSFWorkbook hssfWorkbook) {
        byte[] bytes = new byte[0];
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        try {
            hssfWorkbook.write(byteArrayOutputStream);
            bytes = byteArrayOutputStream.toByteArray();
        } catch (IOException e) {
            log.error("处理Excel文件失败", e);
        }
        return bytes;
    }

    /**
     * 自动根据文件扩展名，调用对应的读取方法
     *
     * @param file       路径
     * @param ext        扩展名
     * @param sheetIdx   读取的sheet页号,从0开始
     * @param columnName 列名
     * @return 为空时，返回null
     */
    public static List<Map<String, Object>> readExcel(File file, String ext, int sheetIdx, String[] columnName) {
        return readExcel(file, ext, null, sheetIdx, columnName);
    }

    /**
     * 自动根据文件扩展名，调用对应的读取方法
     *
     * @param file       路径
     * @param ext        扩展名
     * @param sheetName  读取的sheet名;为空null时,默认读取第一个sheet页
     * @param columnName 列名
     * @return 为空时，返回null
     */
    public static List<Map<String, Object>> readExcel(File file, String ext, String sheetName, String[] columnName) {
        return readExcel(file, ext, sheetName, 0, columnName);
    }

    /**
     * 自动根据文件扩展名，调用对应的读取方法
     *
     * @param file       路径
     * @param ext        后缀
     * @param sheetName  读取的sheet名;为空null时,按照sheetNum读取
     * @param sheetIdx   sheet页号;
     * @param columnName 列名
     * @return 为空时，返回null
     */
    private static List<Map<String, Object>> readExcel(File file, String ext, String sheetName, int sheetIdx, String[] columnName) {
        try {
            if ("xls".equals(ext)) { // 使用xls方式读取
                return readExcel_xls(file, sheetName, sheetIdx, columnName);
            } else if ("xlsx".equals(ext)) { // 使用xlsx方式读取
                return readExcel_xlsx(file, sheetName, sheetIdx, columnName);
            } else { // 依次尝试xls、xlsx方式读取
                log.info("读取Excel:您要操作的文件非excel扩展名，正在尝试以xls方式读取...");
                try {
                    return readExcel_xls(file, sheetName, sheetIdx, columnName);
                } catch (Exception e1) {
                    log.info("读取Excel:尝试以xls方式读取，结果失败！，正在尝试以xlsx方式读取...");
                    try {
                        return readExcel_xlsx(file, sheetName, sheetIdx, columnName);
                    } catch (Exception e2) {
                        log.info("读取Excel:尝试以xls方式读取，结果失败！\n请您确保您的文件是Excel文件，并且无损，然后再试。");
                        return null;
                    }
                }
            }
        } catch (Exception e) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e.printStackTrace(pw);
            log.info(sw.toString());
            try {
                sw.close();
                pw.close();
            } catch (IOException e1) {
                log.error("异常", e1);
            }
        }
        return null;
    }

    /**
     * 读取xls
     *
     * @param file
     * @param sheetName
     * @param sheetIdx
     * @param columnName
     * @return 结果
     */
    private static List<Map<String, Object>> readExcel_xls(File file, String sheetName, int sheetIdx, String[] columnName) {
        List<Map<String, Object>> resultList;
        try {
            HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(file));
            resultList = readExcel(wb, sheetName, sheetIdx, columnName);
        } catch (Exception e) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e.printStackTrace(pw);
            log.info(sw.toString());
            resultList = null;
            try {
                sw.close();
                pw.close();
            } catch (IOException e1) {
                log.error("异常", e1);
            }
        }
        return resultList;
    }

    /**
     * 读取xlsx
     *
     * @param file
     * @param sheetName
     * @param sheetIdx
     * @param columnName
     * @return 结果
     */
    private static List<Map<String, Object>> readExcel_xlsx(File file, String sheetName, int sheetIdx, String[] columnName) {
        List<Map<String, Object>> resultList;
        try {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
            resultList = readExcel(wb, sheetName, sheetIdx, columnName);
        } catch (Exception e) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e.printStackTrace(pw);
            log.info(sw.toString());
            resultList = null;
            try {
                sw.close();
                pw.close();
            } catch (IOException e1) {
                log.error("异常", e1);
            }
        }
        return resultList;
    }

    /**
     * 通用读取Wookbook
     *
     * @param wb         数据
     * @param sheetName  名称
     * @param sheetIdx   序号
     * @param columnName 列集合
     * @return 数据
     * @throws Exception 异常
     */
    public static List<Map<String, Object>> readExcel(Workbook wb, String sheetName, int sheetIdx, String[] columnName) throws Exception {
        Sheet sheet = null;
        List<Map<String, Object>> resultList = new LinkedList<>();
        sheet = ("".equals(sheetName) || null == sheetName) ? wb.getSheetAt(sheetIdx) : wb.getSheet(sheetName);
        int lastRowNum = sheet.getLastRowNum();
        for (int i = readStartRowPos; i <= lastRowNum; i++) {
            Map<String, Object> map = new HashMap<>();
            Row row = sheet.getRow(i);
            if (row == null) continue;
            for (int j = 0; j < columnName.length; j++) {
                map.put(columnName[j], getCellValue(row.getCell(j)));
            }
            resultList.add(map);
        }
        return resultList;
    }

    /***
     * 读取单元格的值
     */
    private static String getCellValue(Cell cell) throws Exception {
        Object result = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    result = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    result = new DecimalFormat("0").format(cell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    result = cell.getBooleanCellValue();
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    result = cell.getCellFormula();
                    break;
                case Cell.CELL_TYPE_ERROR:
                    result = cell.getErrorCellValue();
                    break;
                case Cell.CELL_TYPE_BLANK:
                    break;
                default:
                    result = cell.toString();
                    break;
            }
        }
        String reString = result.toString();
        Pattern p = Pattern.compile("\\s*|\t|\r|\n");
        Matcher m = p.matcher(reString);
        reString = m.replaceAll("").trim();
        return reString;
    }

    /**
     * 创建Excel
     *
     * @param list            数据源
     * @param columnName      Map中存放的key;生成excel列的顺序
     * @param columnAlignName 生成Excel时显示的列的别名
     * @return 数据
     */
    public static HSSFWorkbook getHSSFWookbook(List<Map<String, Object>> list, String[] columnName, String[] columnAlignName) {
        return getHSSFWookbook(list, columnName, columnAlignName, null);
    }

    /**
     * 创建Excel
     *
     * @param list            List数据源
     * @param columnName      Map中存放的key;生成excel列的顺序
     * @param columnAlignName 生成Excel时显示的列的别名
     * @param sheetName       sheet页的名称
     * @return 数据
     */
    public static HSSFWorkbook getHSSFWookbook(List<Map<String, Object>> list, String[] columnName, String[] columnAlignName, String sheetName) {
        HSSFWorkbook workbook = null;
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;
        int rowNum = 0;
        try {
            workbook = createWorkbook(HSSFWorkbook.class);
            sheet = ("".equals(sheetName) || null == sheetName) ? workbook.createSheet() : workbook.createSheet(sheetName);
            //标题格式
            Font font = workbook.createFont();
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFont(font);
            //生成标题
            sheet.createFreezePane(0, 1);
            row = sheet.createRow(rowNum++);
            for (int i = 0; i < columnAlignName.length; i++) {
                cell = row.createCell(i);
                cell.setCellValue(columnAlignName[i]);
                cell.setCellStyle(cellStyle);
            }
            //生成数据
            for (int i = 0; i < list.size(); i++) {
                Map<String, Object> map = list.get(i);
                row = sheet.createRow(rowNum++);
                for (int j = 0; j < columnName.length; j++) {
                    cell = row.createCell(j);
                    Object tmp = map.get(columnName[j]);
                    cell.setCellValue(tmp == null ? "" : String.valueOf(tmp));
                    //设置为字符串
                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                }
            }
            //调整宽度
            for (int i = 0; i < columnAlignName.length; i++) {
                sheet.autoSizeColumn((short) i);
            }
        } catch (Exception e) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e.printStackTrace(pw);
            log.info(sw.toString());
            workbook = null;
            try {
                sw.close();
                pw.close();
            } catch (IOException e1) {
                log.error("异常", e1);
            }
        }
        return workbook;
    }

    /**
     * 创建Excel
     *
     * @param list            数据源
     * @param columnName      Map中存放的key;生成excel列的顺序
     * @param columnAlignName 生成Excel时显示的列的别名
     * @return 结果
     */
    public static XSSFWorkbook getXSSFWookbook(List<Map<String, Object>> list, String[] columnName, String[] columnAlignName) {
        return getXSSFWookbook(list, columnName, columnAlignName, null);
    }

    /**
     * 创建Excel
     *
     * @param list            数据源
     * @param columnName      Map中存放的key;生成excel列的顺序
     * @param columnAlignName 生成Excel时显示的列的别名
     * @param sheetName       sheet页的名称
     * @return 结果
     */
    public static XSSFWorkbook getXSSFWookbook(List<Map<String, Object>> list, String[] columnName, String[] columnAlignName, String sheetName) {
        XSSFWorkbook workbook = null;
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;
        int rowNum = 0;
        try {
            workbook = createWorkbook(XSSFWorkbook.class);
            sheet = ("".equals(sheetName) || null == sheetName) ? workbook.createSheet() : workbook.createSheet(sheetName);
            //标题格式
            Font font = workbook.createFont();
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFont(font);
            //生成标题
            sheet.createFreezePane(0, 1);
            row = sheet.createRow(rowNum++);
            for (int i = 0; i < columnName.length; i++) {
                cell = row.createCell(i);
                cell.setCellValue(columnAlignName[i]);
                cell.setCellStyle(cellStyle);
            }
            //生成数据
            for (int i = 0; i < list.size(); i++) {
                Map<String, Object> map = list.get(i);
                row = sheet.createRow(rowNum++);
                for (int j = 0; j < columnName.length; j++) {
                    cell = row.createCell(j);
                    cell.setCellValue((String) map.get(columnName[j]));
                }
            }

        } catch (Exception e) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e.printStackTrace(pw);
            log.info(sw.toString());
            workbook = null;
            try {
                sw.close();
                pw.close();
            } catch (IOException e1) {
                log.error("异常", e1);
            }
        }
        return workbook;
    }

    /**
     * 创建Excel
     *
     * @param list            数据源
     * @param columnName      Map中存放的key;生成excel列的顺序
     * @param columnAlignName 生成Excel时显示的列的别名
     * @return 结果
     */
    public static HSSFWorkbook getHSSFWookbookByBean(List<?> list, String[] columnName, String[] columnAlignName) {
        return getHSSFWookbookByBean(list, columnName, columnAlignName, null);
    }

    /**
     * 创建Excel
     *
     * @param list            数据源
     * @param columnName      Map中存放的key;生成excel列的顺序
     * @param columnAlignName 生成Excel时显示的列的别名
     * @param sheetName       sheet页的名称
     * @return 结果
     */
    public static HSSFWorkbook getHSSFWookbookByBean(List<?> list, String[] columnName, String[] columnAlignName, String sheetName) {
        HSSFWorkbook workbook = null;
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;
        int rowNum = 0;
        try {
            workbook = createWorkbook(HSSFWorkbook.class);
            sheet = ("".equals(sheetName) || null == sheetName) ? workbook.createSheet() : workbook.createSheet(sheetName);
            //标题格式
            Font font = workbook.createFont();
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFont(font);
            //生成标题
            sheet.createFreezePane(0, 1);
            row = sheet.createRow(rowNum++);
            for (int i = 0; i < columnName.length; i++) {
                cell = row.createCell(i);
                cell.setCellValue(columnAlignName[i]);
                cell.setCellStyle(cellStyle);
            }
            //生成数据
            for (int i = 0; i < list.size(); i++) {
                Object bean = list.get(i);
                Map<String, Object> fieldMap = new HashMap<String, Object>();
                fieldMap = getFieldByObj(bean);
                row = sheet.createRow(rowNum++);
                for (int j = 0; j < columnName.length; j++) {
                    cell = row.createCell(j);
                    cell.setCellValue((String) fieldMap.get(columnName[j]));
                }
            }

        } catch (Exception e) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e.printStackTrace(pw);
            log.info(sw.toString());
            workbook = null;
            try {
                sw.close();
                pw.close();
            } catch (IOException e1) {
                log.error("异常", e1);
            }
        }
        return workbook;
    }

    /**
     * 创建Excel
     *
     * @param list            数据源
     * @param columnName      Map中存放的key;生成excel列的顺序
     * @param columnAlignName 生成Excel时显示的列的别名
     * @return 结果
     */
    public static XSSFWorkbook getXSSFWookbookByBean(List<?> list, String[] columnName, String[] columnAlignName) {
        return getXSSFWookbookByBean(list, columnName, columnAlignName, null);
    }

    /**
     * 创建Excel
     *
     * @param list            数据源
     * @param columnName      Map中存放的key;生成excel列的顺序
     * @param columnAlignName 生成Excel时显示的列的别名
     * @param sheetName       sheet页的名称
     * @return 结果
     */
    public static XSSFWorkbook getXSSFWookbookByBean(List<?> list, String[] columnName, String[] columnAlignName, String sheetName) {
        XSSFWorkbook workbook = null;
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;
        int rowNum = 0;
        try {
            workbook = createWorkbook(XSSFWorkbook.class);
            sheet = ("".equals(sheetName) || null == sheetName) ? workbook.createSheet() : workbook.createSheet(sheetName);
            //标题格式
            Font font = workbook.createFont();
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFont(font);
            //生成标题
            sheet.createFreezePane(0, 1);
            row = sheet.createRow(rowNum++);
            for (int i = 0; i < columnName.length; i++) {
                cell = row.createCell(i);
                cell.setCellValue(columnAlignName[i]);
                cell.setCellStyle(cellStyle);
            }
            //生成数据
            for (int i = 0; i < list.size(); i++) {
                Object bean = list.get(i);
                Map<String, Object> fieldMap = new HashMap<String, Object>();
                fieldMap = getFieldByObj(bean);
                row = sheet.createRow(rowNum++);
                for (int j = 0; j < columnName.length; j++) {
                    cell = row.createCell(j);
                    cell.setCellValue((String) fieldMap.get(columnName[j]));
                }
            }

        } catch (Exception e) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e.printStackTrace(pw);
            log.info(sw.toString());
            workbook = null;
            try {
                sw.close();
                pw.close();
            } catch (IOException e1) {
                log.error("异常", e1);
            }
        }
        return workbook;
    }

    /**
     * 获取某个bean的属性与属性值
     *
     * @param obj
     * @return 结果
     */
    private static Map<String, Object> getFieldByObj(Object obj) throws Exception {
        Map<String, Object> fieldMap = new HashMap<String, Object>();
        Field[] fields = obj.getClass().getDeclaredFields();
        for (Field field : fields) {
            String key = field.getName();
            String value = getFieldValueByName(key, obj);
            fieldMap.put(key, value);
        }
        return fieldMap;
    }

    /**
     * 获取某个属性的值
     *
     * @param fieldName
     * @param obj
     * @return
     * @throws Exception
     */
    private static String getFieldValueByName(String fieldName, Object obj) throws Exception {
        String firstLetter = fieldName.substring(0, 1).toUpperCase();
        String getter = "get" + firstLetter + fieldName.substring(1);
        Method method = obj.getClass().getMethod(getter);
        Object value = method.invoke(obj, new Object[]{});
        return getStringValue(value);
    }

    private static String getStringValue(Object value) {
        if (value instanceof String)
            return (String) value;
        else if (value instanceof Integer)
            return String.valueOf(value);
        else if (value instanceof Boolean)
            return String.valueOf(value);
        else if (value instanceof Double)
            return String.valueOf(value);
        else if (value instanceof Float)
            return String.valueOf(value);
        else if (value instanceof Byte)
            return String.valueOf(value);
        else if (value instanceof Short)
            return String.valueOf(value);
        else if (value instanceof Long)
            return String.valueOf(value);
        else
            return (String) value;
    }

    public static <T> T createWorkbook(Class<T> clazz) {
        try {
            return clazz.newInstance();
        } catch (InstantiationException e) {
            log.error("实例化失败", e);
        } catch (IllegalAccessException e) {
            log.error("实例化失败", e);
        }
        throw new RuntimeException("实例化失败");
    }

    public static void main(String[] args) throws Exception {
        ExcelTools excelUtil = new ExcelTools();
        //读取Excel文件
        String[] columnName = {"ID", "Name", "Age"};
        List<Map<String, Object>> list = excelUtil.readExcel(new File("D:\\logs\\Test.xlsx"), "xlsx", 0, columnName);
        //打印读取结果
        int i = 0;
        for (Map<String, Object> map : list) {
            System.out.println("第" + (i++) + "个:");
            for (Map.Entry<String, Object> entry : map.entrySet()) {
                System.out.println(" " + entry.getKey() + " " + entry.getValue());
            }
        }
        //生成Excel
        OutputStream out = null;
        try {
            //1.由List<Map>生成
            HSSFWorkbook wb = excelUtil.getHSSFWookbook(list, columnName, columnName, "Test");

            //2.由List<bean>生成
			/*com.ryhmp.demo.vo.DemoVo demo=new com.ryhmp.demo.vo.DemoVo();
			demo.setId("1");demo.setName("11");
			List<com.ryhmp.demo.vo.DemoVo> listVo=new LinkedList<com.ryhmp.demo.vo.DemoVo>();
			listVo.add(demo);
			String [] columnNameBean={"id","name"};
			HSSFWorkbook wb=excelUtil.getHSSFWookbookByBean(listVo, columnNameBean, columnNameBean);*/

            File file = new File("D:\\logs\\TestOut.xls");
            out = new FileOutputStream(file);
            wb.write(out);
            out.flush();
        } catch (Exception e) {
            log.error("异常", e);
        } finally {
            if (out != null) out.close();
        }
    }
}
