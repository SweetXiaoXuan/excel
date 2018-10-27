package com.excel.easyexcel;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.*;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelWriteTest {
    /**
     * 每行数据是List<String>无表头
     *
     * @throws IOException
     */
//    @Test
//    public void writeWithoutHead() throws IOException {
//        try (OutputStream out = new FileOutputStream("withoutHead.xlsx");) {
//            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX, false);
//            Sheet sheet1 = new Sheet(1, 0);
//            sheet1.setSheetName("sheet1");
//            List<List<String>> data = new ArrayList<>();
//            for (int i = 0; i < 100; i++) {
//                List<String> item = new ArrayList<>();
//                item.add("item0" + i);
//                item.add("item1" + i);
//                item.add("item2" + i);
//                data.add(item);
//            }
//            writer.write0(data, sheet1);
//            writer.finish();
//        }
//    }

    /**
     * 每行数据是List<String>有表头 非类
     * @throws IOException
     */
    @Test
    public void writeWithHead() throws IOException {
        try (OutputStream out = new FileOutputStream("withHead.xlsx");) {
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
            Sheet sheet1 = new Sheet(1, 0);
            sheet1.setSheetName("sheet1");
            List<List<String>> data = new ArrayList<>();
            for (int i = 0; i < 10; i++) {
                List<String> item = new ArrayList<>();
                item.add("item0" + i);
                item.add("item1" + i);
                item.add("item2" + i);
                data.add(item);
            }
            List<List<String>> head = new ArrayList<List<String>>();
            List<String> headCoulumn1 = new ArrayList<String>();
            List<String> headCoulumn2 = new ArrayList<String>();
            List<String> headCoulumn3 = new ArrayList<String>();
            headCoulumn1.add("第一列");
            headCoulumn2.add("第二列");
            headCoulumn3.add("第三列");
            head.add(headCoulumn1);
            head.add(headCoulumn2);
            head.add(headCoulumn3);
            Table table = new Table(1);
            table.setHead(head);
            TableStyle style = new TableStyle();

//            style.setTableHeadBackGroundColor(IndexedColors.WHITE);

            style.setTableContentBackGroundColor(IndexedColors.WHITE);
            table.setTableStyle(style);
            writer.write0(data, sheet1, table);
            writer.finish();
        }
    }


    /**
     * 除了上面添加表头的方式，我们还可以使用实体类，为其添加com.alibaba.excel.annotation.ExcelProperty注解来生成表头，实体类数据作为Excel数据
     * @throws IOException
     */
    @Test
    public void writeWithHeadClass() throws IOException {
//        try (OutputStream out = new FileOutputStream("withHead.xlsx");) {
//            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
//            Sheet sheet1 = new Sheet(1, 0, ExcelPropertyIndexModel.class);
//            sheet1.setSheetName("sheet1");
            List<ExcelPropertyIndexModel> data = new ArrayList<>();
            for (int i = 0; i < 10; i++) {
                ExcelPropertyIndexModel item = new ExcelPropertyIndexModel();
                item.setName("name" + i);
                item.setAge("age" + i);
                item.setEmail("email" + i);
                item.setAddress("address" + i);
                item.setSax("sax" + i);
                item.setHeigh("heigh" + i);
                item.setLast("last" + i);
                data.add(item);
            }
            ExcelUtil.export(data, ExcelPropertyIndexModel.class, "withHead.xls", "sheet1");


//            TableStyle style = new TableStyle();
//
////            style.setTableHeadBackGroundColor(IndexedColors.WHITE);
//
//            style.setTableContentBackGroundColor(IndexedColors.WHITE);
//            sheet1.setTableStyle(style);
//            writer.write(data, sheet1);
//            writer.finish();
        }

//
//
//    /**
//     * 如果单行表头表头还不满足需求，没关系，还可以使用多行复杂的表头
//     * @throws IOException
//     */
//    @Test
//    public void writeWithMultiHead() throws IOException {
//        try (OutputStream out = new FileOutputStream("withMultiHead.xlsx");) {
//            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
//            Sheet sheet1 = new Sheet(1, 0, MultiLineHeadExcelModel.class);
//            sheet1.setSheetName("sheet1");
//            List<MultiLineHeadExcelModel> data = new ArrayList<>();
//            for (int i = 0; i < 100; i++) {
//                MultiLineHeadExcelModel item = new MultiLineHeadExcelModel();
//                item.p1 = "p1" + i;
//                item.p2 = "p2" + i;
//                item.p3 = "p3" + i;
//                item.p4 = "p4" + i;
//                item.p5 = "p5" + i;
//                item.p6 = "p6" + i;
//                item.p7 = "p7" + i;
//                item.p8 = "p8" + i;
//                item.p9 = "p9" + i;
//                data.add(item);
//            }
//            writer.write(data, sheet1);
//            writer.finish();
//        }
//    }
//    public static class MultiLineHeadExcelModel extends BaseRowModel {
//
//        @ExcelProperty(value = { "表头1", "表头1", "表头31" }, index = 0)
//        private String p1;
//
//        @ExcelProperty(value = { "表头1", "表头1", "表头32" }, index = 1)
//        private String p2;
//
//        @ExcelProperty(value = { "表头3", "表头3", "表头3" }, index = 2)
//        private String p3;
//
//        @ExcelProperty(value = { "表头4", "表头4", "表头4" }, index = 3)
//        private String p4;
//
//        @ExcelProperty(value = { "表头5", "表头51", "表头52" }, index = 4)
//        private String p5;
//
//        @ExcelProperty(value = { "表头6", "表头61", "表头611" }, index = 5)
//        private String p6;
//
//        @ExcelProperty(value = { "表头6", "表头61", "表头612" }, index = 6)
//        private String p7;
//
//        @ExcelProperty(value = { "表头6", "表头62", "表头621" }, index = 7)
//        private String p8;
//
//        @ExcelProperty(value = { "表头6", "表头62", "表头622" }, index = 8)
//        private String p9;
//
//        public String getP1() {
//            return p1;
//        }
//
//        public void setP1(String p1) {
//            this.p1 = p1;
//        }
//
//        public String getP2() {
//            return p2;
//        }
//
//        public void setP2(String p2) {
//            this.p2 = p2;
//        }
//
//        public String getP3() {
//            return p3;
//        }
//
//        public void setP3(String p3) {
//            this.p3 = p3;
//        }
//
//        public String getP4() {
//            return p4;
//        }
//
//        public void setP4(String p4) {
//            this.p4 = p4;
//        }
//
//        public String getP5() {
//            return p5;
//        }
//
//        public void setP5(String p5) {
//            this.p5 = p5;
//        }
//
//        public String getP6() {
//            return p6;
//        }
//
//        public void setP6(String p6) {
//            this.p6 = p6;
//        }
//
//        public String getP7() {
//            return p7;
//        }
//
//        public void setP7(String p7) {
//            this.p7 = p7;
//        }
//
//        public String getP8() {
//            return p8;
//        }
//
//        public void setP8(String p8) {
//            this.p8 = p8;
//        }
//
//        public String getP9() {
//            return p9;
//        }
//
//        public void setP9(String p9) {
//            this.p9 = p9;
//        }
//    }
}
