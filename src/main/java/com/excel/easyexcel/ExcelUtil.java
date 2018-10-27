package com.excel.easyexcel;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.ExcelHeadProperty;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.metadata.TableStyle;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.context.GenerateContextImpl;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtil {


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



    public static void export1(List data, Class clazz, String fileName, String sheetName) throws IOException {
        try (OutputStream out = new FileOutputStream(fileName)) {
            Sheet sheet1 = new Sheet(1, 0, clazz);
            sheet1.setSheetName(sheetName);
            TableStyle style = new TableStyle();
            style.setTableContentBackGroundColor(IndexedColors.WHITE);
            sheet1.setTableStyle(style);

            GenerateContextImpl context = new GenerateContextImpl(out, ExcelTypeEnum.XLS, true);
            Table table = new Table(1);
            table.setTableStyle(style);
            table.setTableNo(1);
            table.setClazz(ExcelPropertyIndexModel.class);
            table.setHead(getHead(null));

            Workbook wb = new XSSFWorkbook();
            org.apache.poi.ss.usermodel.Sheet sheet = wb.createSheet();
            context.setCurrentSheet(sheet);
            context.buildCurrentSheet(sheet1);
            context.buildTable(table);
        }
    }

    private static List<List<String>> getHead(List<String> headName) {
        headName = new ArrayList<>();
        headName.add("1");
        headName.add("2");
        headName.add("3");
        List<List<String>> head = new ArrayList<List<String>>();
        for (String str : headName) {
            List<String> headCoulumn = new ArrayList<String>();
            headCoulumn.add(str);
            head.add(headCoulumn);
        }
        return head;
    }




//    public static void main(String[] args) throws Exception {
//        //准备excel输出流
//        OutputStream ops = new FileOutputStream("C:/Users/Administrator/Desktop/exportStudent.xlsx");
//        //创建excel上下文实例,它的构成需要配置文件的路径
//        Excel
//        ExcelContext context = new ExcelContext("excel-config.xml");
//        //获取POI创建结果
//        Workbook workbook = context.createExcel("student",getStudents());
//        workbook.write(ops);
//        ops.close();
//        workbook.close();
//    }

    //获取模拟数据,数据库数据...
    public static List<StudentModel> getStudents(){
        int size = 10;
        List<StudentModel> students = new ArrayList<>(size);
        for(int i=0;i<size;i++){
            StudentModel stu = new StudentModel();
            stu.setName("张三"+i);
            stu.setAge(20+i);
            stu.setStudentNo("Stu_"+i);
            stu.setStatus(i%2==0?1:0);
            stu.setCreateUser("王五"+i);

            //创建复杂对象
            if(i % 2==0){
                BookModel book = new BookModel();
                book.setBookName("Thinking in java");
                AuthorModel author = new AuthorModel();
                author.setAuthorName("Bruce Eckel");
                book.setAuthor(author);
                stu.setBookModel(book);
            }

            students.add(stu);
        }
        return students;
    }

}
