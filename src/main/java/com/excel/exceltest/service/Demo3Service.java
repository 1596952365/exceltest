package com.zx.excelImport.service;

import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.excel.exceltest.util.XxlsAbstract;
import org.springframework.stereotype.Service;

import com.excel.exceltest.util.XxlsAbstract;

@Service
public class Demo3Service extends XxlsAbstract {

    int sheet = 1;
    public Demo3Service(){

    }
    public Demo3Service(int sheet){
        this.sheet = sheet;
    }
    static List<Map<String, Object>> mapList = new ArrayList<Map<String, Object>>();
    private static final String map1Name[] = {"STUDENTNAME","STUDENTNO","STUDENTSEX","BIRTHDAY","CARD"};
    private static final String map2Name[] = {"PARENTNAME","TEL","RELATION"};
    List<String> map1Name = new ArrayList();

    @Override
    public void optRows(int sheetIndex, int curRow, List<String> rowlist) throws SQLException {
        if(null!=rowlist&&rowlist.size()>0){
            Map<String, Object> map = new HashMap<String, Object>();
            if(curRow>=1){
                if(sheet==1){
                    for(int i=0;i<map1Name.length;i++){
                        if(i<rowlist.size()){
                            map.put(map1Name[i], rowlist.get(i));
                        }
                    }
                }else if(sheet==2){
                    for(int i=0;i<map2Name.length;i++){
                        if(i<rowlist.size()){
                            map.put(map2Name[i], rowlist.get(i));
                        }
                    }
                }
                mapList.add(map);
            }
        }

    }

    public void readExcel(String xlsUlr) throws Exception{
        mapList = new ArrayList<Map<String,Object>>();
        Demo3Service  ds = new Demo3Service();
        ds.processOneSheet(xlsUlr, 1);
        for(int i=0;i<mapList.size();i++){
            Map<String,Object> map = mapList.get(i);
            System.out.println(map.get("STUDENTNAME"));
            System.out.println(map.get("STUDENTNO"));
            System.out.println(map.get("STUDENTSEX"));
            System.out.println(map.get("BIRTHDAY"));
            System.out.println(map.get("CARD"));
        }

        mapList.clear();
        Demo3Service  ds2 = new Demo3Service(2);
        ds2.processOneSheet(xlsUlr, 2);
        for(int i=0;i<mapList.size();i++){
            Map<String,Object> map = mapList.get(i);
            System.out.println(map.get("PARENTNAME"));
            System.out.println(map.get("TEL"));
            System.out.println(map.get("RELATION"));
        }
    }
}
