package com.workpilot.dto;

import lombok.Data;

@Data
public class ReportDTO {

    /**
     * 日期
     * Excel导入/导出对应列名：日期
     * 日期格式统一指定，避免解析异常
     */
    private String reportDate;

    /**
     * 星期
     * Excel导入/导出对应列名：星期
     */
    private String weekDay;

    /**
     * 有效工时
     * Excel导入/导出对应列名：有效工时
     * 用BigDecimal避免浮点精度丢失，适配Excel小数格式
     */
    private String workHours;

    /**
     * 工作内容
     * Excel导入/导出对应列名：工作内容
     * 长文本自动换行，列宽适配
     */
    private String workContent;
}
