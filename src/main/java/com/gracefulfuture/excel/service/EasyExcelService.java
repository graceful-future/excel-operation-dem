package com.gracefulfuture.excel.service;

import cn.hutool.core.bean.BeanUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.gracefulfuture.excel.entity.Student;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.util.ObjectUtils;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @author chenkun
 * @version 1.0
 * @description easyexcel服务接口
 * @create 2021/7/23 10:29
 */
public class EasyExcelService {
    @Value("${file.target.path}")
    private String targetPath;

    public static void writeExcel(String filePath,String sheetName,String[] heads,List datas) throws FileNotFoundException {
//        List<Student> dataList = getDataList();
//        EasyExcel.write(filePath,Student.class).sheet("学生列表").doWrite(dataList);

        OutputStream outputStream = new FileOutputStream(filePath);
        ExcelWriter excelWriter = new ExcelWriter(outputStream, ExcelTypeEnum.XLSX);
        Sheet sheet = new Sheet(1,0);
        sheet.setSheetName(sheetName);
        Table table = new Table(1);
        List<List<String>> tableTitles = new ArrayList();

        for(int i = 0; i < heads.length; ++i) {
            tableTitles.add(Arrays.asList(heads[i]));
        }

        table.setHead(tableTitles);
        List<List<String>> dataList = new ArrayList();
        Iterator datasIterator = datas.iterator();

        while(datasIterator.hasNext()) {
            Map<String, Object> data = (Map)datasIterator.next();
            List<String> list = new ArrayList();
            Iterator dataIterator = data.entrySet().iterator();

            while(dataIterator.hasNext()) {
                Map.Entry<String, Object> m = (Map.Entry)dataIterator.next();
                list.add(!ObjectUtils.isEmpty(m.getValue()) ? m.getValue().toString() : "");
            }

            dataList.add(list);
        }

        excelWriter.write0(dataList, sheet, table);
        excelWriter.finish();
    }

    public static void readExcel(String filePath){
        EasyExcel.read(filePath,Student.class,new ExcelReadListener()).sheet().doRead();
    }

    //创建方法返回list集合
    private static List<Student> getDataList() {
        List<Student> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            Student student = new Student();
            student.setNumber("1000" + i);
            student.setFullName("Tom" + (i + 1));
            student.setAge(10 + (i % 10));
            student.setAddress("金阳南路");
            student.setClassNo("2021000" + (i + 1));
            list.add(student);
        }
        return list;
    }

    public static void main(String[] args) throws FileNotFoundException {
        String filePath = "C:\\Users\\Administrator\\Desktop\\学生数据.xlsx";
//        String sheetName = "学生列表";
//        String[] titles = new String[]{"学号","姓名","年龄","班号","住址"};
//        List<Student> studentList = getDataList();
//        List<LinkedHashMap> datas = new ArrayList<>();
//        AtomicInteger order = new AtomicInteger(1);
//        studentList.stream().forEach(student -> {
//            LinkedHashMap<String,Object> linkedHashMap = new LinkedHashMap();
//            linkedHashMap = (LinkedHashMap<String, Object>) BeanUtil.beanToMap(student);
////            linkedHashMap.put("order", order.getAndIncrement());
//            datas.add(linkedHashMap);
//        });
//        writeExcel(filePath,sheetName,titles,datas);
        readExcel(filePath);
    }
}
