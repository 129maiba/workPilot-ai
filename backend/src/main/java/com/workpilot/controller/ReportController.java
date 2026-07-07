package com.workpilot.controller;

import com.workpilot.dto.ReportDTO;
import com.workpilot.service.ReportService;
import com.workpilot.util.Result;
import jakarta.annotation.Resource;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;


@RestController
@RequestMapping("/api/report")
public class ReportController {
    @Resource
    ReportService reportService;

    @PostMapping("/upload")
    public Result upload(@RequestParam("excelFile") MultipartFile file) {
        List<ReportDTO> reportDTOList = reportService.uploadReport(file);
        if(reportDTOList != null && reportDTOList.size()>=0){
            return Result.success(reportDTOList);
        }else{
            return Result.fail("上传文件失败");
        }
    }

}
