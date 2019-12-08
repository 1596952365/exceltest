package com.excel.exceltest.Control;

import com.excel.exceltest.util.ImportUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.FileNameMap;
import java.util.Collection;

@Controller
public class test {

    @RequestMapping("demo1")
    public String demo1(){

        return "demo2";

    }

    /*
    * 下载模板访问
    * */
    @RequestMapping("downloadTemplate")
    @ResponseBody
    public void downloadTemplate(HttpServletResponse response){

        ImportUtil excel = new ImportUtil();
        String FileName ="7月份南京宜家运输结算报表.xls";
        String FilePath = "G:/Excel/src/main/resources/templates/";
        try {
            excel.downFile(FileName,FilePath,response);
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    /*上传excel单个上传
    *
    * */
    @RequestMapping("importFile")
    public String importFile(HttpServletRequest request){
        ImportUtil excel = new ImportUtil();  //加载工具类
        Workbook workbook = excel.importFile(request);  //获取到整个excel的工作表
        Sheet sheet = workbook.getSheetAt(0);//获取第一张工作表
        //开始遍历整张工作表
        for (int i = 0;;i++){
            //得到第一行的数据  如果没有表头就跳出循环
            Row row = sheet.getRow(i);
            //如果这一行为空那么就跳出循环
            if (row==null){
                break;
            }
            //获取导入的表的列数
            int num = row.getPhysicalNumberOfCells();
            System.out.println(num+"+++++++++++++++++++++++");
            //判断每一行是否为空 为空的话就直接循环遍历下一行
            if (excel.isBlankRow(row)){
                    continue;
            }
            //循环判断每一列的数据
            for (int j = 0;j<=num;j++){
                System.out.println(row.getCell(j));
            }
        }
        return "success";
    }
    /*上传excel多个上传
     *
     * */
    @RequestMapping("importFile2")
    public String importFile2(HttpServletRequest request){
        ImportUtil excel = new ImportUtil();
        Collection<MultipartFile> multipartFiles = excel.importFile2(request);
        for (MultipartFile myFiles:multipartFiles) {
            Workbook workbook = excel.getWorkbook(myFiles);
            Sheet sheetAt = workbook.getSheetAt(0);
            for (int j = 0 ;;j++){
                Row row = sheetAt.getRow(j);
                if (row==null){
                break;
                }

                int num = row.getPhysicalNumberOfCells();
                System.out.println(num+"+++++++++++++++++++++++");
                //判断每一行是否为空 为空的话就直接循环遍历下一行
                if (excel.isBlankRow(row)){
                    continue;
                }
                //循环判断每一列的数据
                for (int x = 0;x<=num;x++){
                    System.out.println(row.getCell(x));
                }
            }
        }

        return "success";
    }

    /**
     *一个excel多个表读取
     * */
    @RequestMapping("importFile3")
    public String importFile3(HttpServletRequest request,MultipartFile file){

        return "success";
    }


}
