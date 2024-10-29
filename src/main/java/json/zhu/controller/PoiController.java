package json.zhu.controller;

import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import json.zhu.ExcelUtils;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Controller;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Controller
public class PoiController {

    @PostMapping(value = "/exportExcel")
    public void exportExcel(HttpServletResponse response) throws IOException {

        TemplateExportParams templatePath = new TemplateExportParams(ResourceUtils.getFile("classpath:template/poi-hashMap.xlsx").getAbsolutePath(),false,"Sheet1");
        Map<String, Object> map = new HashMap<>();
        map.put("className", "三年级");
        map.put("teacher", "王老师");

        List<Map<String,Object>> list = new ArrayList<>();
        Map<String, Object> map1 = new HashMap<>();
        map1.put("id",1);
        map1.put("name","张三");
        map1.put("age",21);
        Map<String, Object> map2 = new HashMap<>();
        map2.put("id",2);
        map2.put("name","李四");
        map2.put("age",22);
        list.add(map1);
        list.add(map2);
        map.put("studentList",list);

        String fileName = "我的excel";

        ExcelUtils.exportExcel(templatePath, map, fileName, response);

    }
}
