package com.example.demo.controller;

import com.example.demo.entity.ExcelTitle;
import com.example.demo.service.ExcelFileExportUtils;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.util.*;

@Controller
public class testController {

    @RequestMapping("/hello")
    @ResponseBody
    public String hello() {
        return "hello world";
    }


    @RequestMapping(value = "/exportExcel", method = RequestMethod.GET)
    public void exportExcel(HttpServletResponse resp) {
        LinkedHashMap<String, ExcelTitle> title = new LinkedHashMap<>();
        title.put("title1", new ExcelTitle("标题1", HorizontalAlignment.LEFT));
        title.put("title2", new ExcelTitle("标题2", HorizontalAlignment.CENTER));
        title.put("title3", new ExcelTitle("标题3", HorizontalAlignment.CENTER));
        title.put("title4", new ExcelTitle("标题4", HorizontalAlignment.CENTER));
        title.put("title5", new ExcelTitle("标题5", HorizontalAlignment.RIGHT));

        List<Map<String, Object>> contents = new ArrayList<>();
        for (int i = 0; i < 30000; i++) {
            Map<String, Object> theCont = new HashMap<>();
            theCont.put("title1", "内容1-" + i);
            theCont.put("title2", "内容2-" + i);
            theCont.put("title3", "内容3-" + i);
            theCont.put("title4", "内容4-" + i);
            theCont.put("title5", "内容5-" + i);
            contents.add(theCont);
        }

        String filename = "导出数据1";

        ExcelFileExportUtils.exportExcelFile(title, contents, filename, false, resp);
    }

}
