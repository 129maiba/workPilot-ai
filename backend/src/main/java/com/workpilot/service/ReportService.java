package com.workpilot.service;

import com.workpilot.dto.ReportDTO;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;


public interface ReportService {

    public List<ReportDTO> uploadReport(MultipartFile file);
}
