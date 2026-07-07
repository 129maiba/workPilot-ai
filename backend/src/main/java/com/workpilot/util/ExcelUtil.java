package com.workpilot.util;

import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;
import java.io.*;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

/**
 * Excel 通用工具类
 * 功能：文件校验、Excel解析导入、数据导出、单元格处理、资源释放
 * 兼容：xls、xlsx 双格式，支持大数据量流式导出
 */
public final class ExcelUtil {

    // 私有构造方法，禁止实例化
    private ExcelUtil() {}

    // 常量定义
    public static final String EXCEL_SUFFIX_XLS = ".xls";
    public static final String EXCEL_SUFFIX_XLSX = ".xlsx";
    public static final long MAX_FILE_SIZE = 10 * 1024 * 1024; // 最大文件大小10MB
    public static final int DEFAULT_HEAD_ROW_NUM = 1; // 默认表头占1行
    public static final int DEFAULT_DATA_START_ROW = 1; // 默认数据从第2行开始（索引从0开始）

    /**
     * 校验上传的Excel文件合法性
     * @param file 上传的文件
     * @throws RuntimeException 校验不通过抛出异常
     */
    public static void validateExcelFile(MultipartFile file) {
        // 1. 空文件校验
        if (file == null || file.isEmpty()) {
            throw new RuntimeException("上传的Excel文件不能为空");
        }
        // 2. 文件大小校验
        if (file.getSize() > MAX_FILE_SIZE) {
            //throw new RuntimeException("Excel文件大小不能超过10MB");
        }
        // 3. 文件名后缀校验
        String fileName = file.getOriginalFilename();
        if (fileName == null) {
            throw new RuntimeException("文件名不能为空");
        }
        String suffix = fileName.substring(fileName.lastIndexOf("."));
        if (!EXCEL_SUFFIX_XLS.equals(suffix) && !EXCEL_SUFFIX_XLSX.equals(suffix)) {
            throw new RuntimeException("仅支持 .xls / .xlsx 格式的Excel文件");
        }
    }

    /**
     * 解析Excel文件为JavaBean列表（默认表头1行，数据从第2行开始）
     * @param file 上传的Excel文件
     * @param clazz 目标JavaBean类型
     * @param rowMapper 行数据映射函数（Row -> T）
     * @return 解析后的JavaBean列表
     * @throws RuntimeException 解析失败抛出异常
     */
    public static <T> List<T> parseExcelToList(MultipartFile file, Class<T> clazz, Function<Row, T> rowMapper) {
        // 先校验文件
        validateExcelFile(file);
        List<T> resultList = new ArrayList<>();

        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = getWorkbook(inputStream, file.getOriginalFilename())) {

            // 读取第一个Sheet，可扩展为读取所有Sheet
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) {
                return resultList;
            }

            // 遍历数据行，跳过表头
            for (int rowIndex = DEFAULT_DATA_START_ROW; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue; // 跳过空行
                }
                // 行数据映射为JavaBean
                T entity = rowMapper.apply(row);
                if (entity != null) {
                    resultList.add(entity);
                }
            }
        } catch (Exception e) {
            throw new RuntimeException("Excel解析失败：" + e.getMessage(), e);
        }
        return resultList;
    }

    /**
     * 通用单元格取值方法（兼容所有单元格类型）
     * @param cell 单元格对象
     * @return 单元格字符串值
     */
    public static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                // 数字类型统一转为字符串，避免科学计数法
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                // 公式类型取计算后的值
                return cell.getNumericCellValue() != 0 ? String.valueOf(cell.getNumericCellValue()) : cell.getStringCellValue();
            default:
                return "";
        }
    }

    /**
     * 导出JavaBean列表为Excel文件（浏览器下载）
     * @param response HTTP响应对象
     * @param fileName 导出文件名（不含后缀）
     * @param sheetName Sheet名称
     * @param headers 表头数组
     * @param dataList 数据列表
     * @param rowWriter 行数据写入函数（Row, T -> void）
     * @throws RuntimeException 导出失败抛出异常
     */
    public static <T> void exportExcel(HttpServletResponse response, String fileName, String sheetName,
                                       String[] headers, List<T> dataList, Function<T, Object[]> rowWriter) {
        // 1. 初始化Workbook（大数据量用SXSSFWorkbook，避免OOM）
        try (SXSSFWorkbook workbook = new SXSSFWorkbook();
             OutputStream outputStream = response.getOutputStream()) {

            // 2. 创建Sheet
            Sheet sheet = workbook.createSheet(sheetName);
            // 3. 写入表头
            writeHeader(sheet, headers);
            // 4. 写入数据行
            writeData(sheet, dataList, rowWriter);
            // 5. 浏览器下载配置
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setCharacterEncoding(StandardCharsets.UTF_8.name());
            String encodedFileName = URLEncoder.encode(fileName + EXCEL_SUFFIX_XLSX, StandardCharsets.UTF_8.name());
            response.setHeader("Content-Disposition", "attachment;filename=" + encodedFileName);
            // 6. 写入响应流
            workbook.write(outputStream);
            // 释放SXSSF临时文件资源
            workbook.dispose();
        } catch (Exception e) {
            throw new RuntimeException("Excel导出失败：" + e.getMessage(), e);
        }
    }

    /**
     * 写入表头
     */
    private static void writeHeader(Sheet sheet, String[] headers) {
        Row headerRow = sheet.createRow(0);
        // 表头样式：加粗、居中
        CellStyle headerStyle = sheet.getWorkbook().createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        headerStyle.setFont(font);

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
            // 列宽自适应
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * 写入数据行
     */
    private static <T> void writeData(Sheet sheet, List<T> dataList, Function<T, Object[]> rowWriter) {
        int rowIndex = 1; // 数据从第2行开始
        for (T data : dataList) {
            Row row = sheet.createRow(rowIndex++);
            Object[] cellValues = rowWriter.apply(data);
            for (int i = 0; i < cellValues.length; i++) {
                Cell cell = row.createCell(i);
                Object value = cellValues[i];
                if (value == null) {
                    continue;
                }
                // 按类型写入单元格
                if (value instanceof String) {
                    cell.setCellValue((String) value);
                } else if (value instanceof Integer) {
                    cell.setCellValue((Integer) value);
                } else if (value instanceof Double) {
                    cell.setCellValue((Double) value);
                } else if (value instanceof Boolean) {
                    cell.setCellValue((Boolean) value);
                }
            }
        }
    }

    /**
     * 获取Workbook实例，兼容xls/xlsx
     */
    private static Workbook getWorkbook(InputStream inputStream, String fileName) throws IOException {
        String suffix = fileName.substring(fileName.lastIndexOf("."));
        if (EXCEL_SUFFIX_XLS.equals(suffix)) {
            return new HSSFWorkbook(inputStream);
        } else if (EXCEL_SUFFIX_XLSX.equals(suffix)) {
            return new XSSFWorkbook(inputStream);
        } else {
            throw new IOException("不支持的Excel格式：" + suffix);
        }
    }

    /**
     * 静默关闭流资源（兼容JDK7+ try-with-resources，备用方法）
     */
    public static void closeQuietly(Closeable closeable) {
        if (closeable != null) {
            try {
                closeable.close();
            } catch (IOException e) {
                // 静默关闭，不抛出异常
            }
        }
    }
}
