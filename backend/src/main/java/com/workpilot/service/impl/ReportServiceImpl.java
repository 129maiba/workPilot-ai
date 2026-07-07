package com.workpilot.service.impl;

import com.workpilot.dto.ReportDTO;
import com.workpilot.service.ReportService;
import com.workpilot.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;


@Service
public class ReportServiceImpl implements ReportService {

    @Override
    public List<ReportDTO> uploadReport(MultipartFile file) {
        List<ReportDTO> reportDTOList  = new ArrayList<>();
        // 1. 判断文件是否为空
        if (file.isEmpty()) {
            new Exception("文件不能为空");
        }
        // 2. 校验后缀
        String fileName = file.getOriginalFilename();
        if (!fileName.endsWith(".xlsx") && !fileName.endsWith(".xls")) {
            new Exception("仅支持xls/xlsx格式");
        }
        try {
            //直接读取Excel数据（POI解析）
            // List<ExcelData> list = ExcelUtil.readExcel(file.getInputStream());
            // 1. 定义数据行映射函数
            Function<Row, ReportDTO> rowMapper = row -> {
                ReportDTO dto = new ReportDTO();
                dto.setReportDate(ExcelUtil.getCellValue(row.getCell(0)));
                dto.setWeekDay(ExcelUtil.getCellValue(row.getCell(1)));
                dto.setWorkHours(ExcelUtil.getCellValue(row.getCell(2)));
                dto.setWorkContent(ExcelUtil.getCellValue(row.getCell(3)));
                return dto;
            };
            reportDTOList  =  ExcelUtil.parseExcelToList(file, ReportDTO.class,rowMapper);

        } catch (Exception e) {
            new Exception(e.getMessage());
        }
        return reportDTOList;
    }
}
