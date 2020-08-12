package com.ghb.excel;

import com.alibaba.excel.EasyExcel;

public class ExcelDemo {
    public static void main(String[] args) {
        String OutfileName ="C:\\Users\\ghb\\Desktop\\原始文件\\SG.xlsx";
        EasyExcel.read(OutfileName,WriteModel.class, new ExcelListener()).sheet().doRead();
    }
}
