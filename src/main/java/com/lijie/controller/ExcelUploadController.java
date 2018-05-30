package com.lijie.controller;

import com.alibaba.fastjson.JSONObject;
import com.google.common.base.Strings;
import com.lijie.entity.BckMes;
import com.lijie.excelFile.ExcelFile;
import com.lijie.service.ExcelUploadService;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.Objects;

@Controller
@RequestMapping("/upload")
public class ExcelUploadController {

    @Autowired
    private ExcelUploadService excelUploadService;

    @RequestMapping("/fileToJson")
    @ResponseBody
    public String changeFile(@RequestParam String fileType, @RequestParam MultipartFile file,@RequestParam int rowNum ){
        Objects.requireNonNull(Strings.isNullOrEmpty(fileType)?null:fileType,"文件类型不能为空");
        Objects.requireNonNull(file,"excel文件不能为空");
        rowNum = rowNum>1?rowNum:1;
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
            ExcelFile excelFile = new ExcelFile(fileType);
            Object data = excelUploadService.changeFile(workbook,excelFile,rowNum);
            return BckMes.successMes(data);
        } catch (IOException e) {
            return BckMes.errorMes(e.toString());
        }

    }

}
