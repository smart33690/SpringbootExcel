package com.xiaopotian.Controller;

import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by zouLu on 2017-12-14.
 */
@RestController
public class ExcelController {
    @RequestMapping(value = "/excel", method = RequestMethod.GET)
    public void excel(HttpServletResponse response) throws Exception {
        ExcelData data = new ExcelData();
        data.setName("hello");
        List<String> titles = new ArrayList();
        titles.add("项目阶段");
        titles.add("区域");
        titles.add("公司");
        data.setTitles(titles);

        List<List<Object>> rows = new ArrayList();
        List<Object> row = new ArrayList();
        row.add("拿地方案");
        row.add("项目启动会");
        row.add("节点进展会议");
        row.add("胜利大会");
        rows.add(row);

        data.setRows(rows);


        //生成本地

        ExportExcelUtils.exportExcel(response,"hello.xlsx",data);
    }
}
