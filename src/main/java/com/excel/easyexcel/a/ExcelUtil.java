package com.excel.easyexcel.a;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.TableStyle;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.excel.easyexcel.ImportInfo;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

@SuppressWarnings("all")
public class ExcelUtil {
    @Test
    public void testWriter() throws FileNotFoundException {
        OutputStream out = new FileOutputStream("test.xls");
        try {
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLS);
            Sheet sheet = new Sheet(0,0, ImportInfo.class);
            TableStyle style = new TableStyle();
            style.setTableContentBackGroundColor(IndexedColors.WHITE);
            sheet.setTableStyle(style);
            sheet.setHead(getHead(null));
            writer.write(getDate(),sheet);
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static void export(List data, Class clazz, String fileName, String sheetName) throws IOException {
        try (OutputStream out = new FileOutputStream(fileName)) {
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
            Sheet sheet1 = new Sheet(1, 0, clazz);
            sheet1.setSheetName(sheetName);
            TableStyle style = new TableStyle();
            style.setTableContentBackGroundColor(IndexedColors.WHITE);
            sheet1.setTableStyle(style);
            writer.write(data, sheet1);
            writer.finish();

        }
    }

    private static List<List<String>> getHead(List<String> headName) {
        headName = new ArrayList<>();
        headName.add("第一列");
        headName.add("第二列");
        headName.add("第三列");
        List<List<String>> head = new ArrayList<List<String>>();
        for (String str : headName) {
            List<String> headCoulumn = new ArrayList<String>();
            headCoulumn.add(str);
            head.add(headCoulumn);
        }
        return head;
    }

    public List<ImportInfo> getDate(){
        List<ImportInfo> list = new ArrayList<ImportInfo>();
        ImportInfo info = new ImportInfo();
        info.setAge(12);
        info.setName("zhangsan");
        info.setEmail("11111@qq.com");
        ImportInfo info1 = new ImportInfo();
        info1.setAge(12);
        info1.setName("zhangsan1");
        info1.setEmail("11111@qq.com");
        ImportInfo info2 = new ImportInfo();
        info2.setAge(12);
        info2.setName("zhangsan2");
        info2.setEmail("11111@qq.com");
        list.add(info);
        list.add(info1);
        list.add(info2);
        return list;
    }
}
