package com.lijie.controller;

import com.alibaba.fastjson.JSONObject;
import com.lijie.excelFile.ExcelFile;
import com.lijie.service.ExcelDowService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import java.io.File;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author lijie7
 * @date 2018/5/25
 * @Description
 * @modified By
 */
@RestController(value = "/")
public class ExcelDowController {
    @Autowired
    private ExcelDowService excelDowService;

    @RequestMapping(value = "/downloadFile", method = RequestMethod.POST)
    @ResponseBody
    public ResponseEntity<byte[]> down(@RequestBody JSONObject jsonObject) {
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", "dict.txt");
        headers.setCacheControl("no-cache, no-store, must-revalidate");
        headers.setContentDispositionFormData("excel", System.currentTimeMillis() + ".xls");
        ExcelFile file = new ExcelFile(jsonObject.getString("type"));
        ByteArrayOutputStream byteArrayOutputStream = excelDowService.exportFile(file, jsonObject);

        byte[] body = byteArrayOutputStream.toByteArray();
        try {
            File file1 = new File("12NEW.xlsx");
            FileOutputStream fileOutputStream = new FileOutputStream(file1);
            fileOutputStream.write(byteArrayOutputStream.toByteArray());
            fileOutputStream.flush();
            fileOutputStream.close();
            byteArrayOutputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return ResponseEntity.ok().headers(headers).body(body);

    }
}
