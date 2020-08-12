package com.ghb.excel;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.*;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.FillPatternType;

@Data
@NoArgsConstructor
@AllArgsConstructor
@ContentRowHeight(20)
@HeadRowHeight(20)
@ColumnWidth(25)
@Builder
public class WriteModel {

    @ExcelProperty(value = "入库时间", index = 0)
    private String storageTime;

    @ExcelProperty(value = "主客户编码", index = 1)
    private String custCode;

    @ExcelProperty(value = "主客户", index = 2)
    private String custName;
    @ExcelProperty(value = "分客户编码", index = 3)
    private String subCustCode;
    @ExcelProperty(value = "分客户", index = 4)
    private String subCustName;

    @ExcelProperty(value = "进仓编号", index = 5)
    private String custId;

    @ExcelProperty(value = "PO号", index = 6)
    private String poNo;

    @ExcelProperty(value = "SKU号", index = 7)
    private String skuNo;

    @ExcelProperty(value = "ITEM", index = 8)
    private String item;

    @ExcelProperty(value = "唛头", index = 9)
    private String markNo;

    @ExcelProperty(value = "款号", index = 10)
    private String typeNo;

    @ExcelProperty(value = "色号", index = 11)
    private String colorNo;

    @ExcelProperty(value = "尺码", index = 12)
    private String sizeNo;

    @ExcelProperty(value = "入库单号", index = 13)
    private String orderNo;

    @ExcelProperty(value = "货物类型", index = 14)
    private String custType;

    @ExcelProperty(value = "批次号", index = 15)
    private String batchNo;

    @ExcelProperty(value = "批次属性08", index = 16)
    private String attribute08;

    @ExcelProperty(value = "批次属性09", index = 17)
    private String attribute09;

    @ExcelProperty(value = "批次属性10", index = 18)
    private String attribute10;

    @ExcelProperty(value = "批次属性11", index = 19)
    private String attribute11;

    @ExcelProperty(value = "批次属性12", index = 20)
    private String attribute12;

    @ExcelProperty(value = "供应商编码", index = 21)
    private String supplierCode;

    @ExcelProperty(value = "供应商名称", index = 22)
    private String supplierName;

    @ExcelProperty(value = "托盘号", index = 23)
    private String LPNCode;

    @ExcelProperty(value = "现有数量", index = 24)
    private String qty_on_hand;

    @ExcelProperty(value = "长", index = 25)
    private String length;

    @ExcelProperty(value = "宽", index = 26)
    private String width;

    @ExcelProperty(value = "高", index = 27)
    private String high;

    @ExcelProperty(value = "重量", index = 28)
    private String weight;
    @ExcelProperty(value = "货位", index = 29)
    private String loactionCode;

    @ExcelProperty(value = "商品品质", index = 30)
    private String quality;

    @ExcelProperty(value = "库存状态", index = 31)
    private String status;

    @ExcelProperty(value = "商品编码", index = 32)
    private String productCode;

    @ExcelProperty(value = "商品条码", index = 33)
    private String barCode;

    @ExcelProperty(value = "货主商品编码", index = 34)
    private String custProductCode;

    @ExcelProperty(value = "商品名称", index = 35)
    private String commodityName;

    @ExcelProperty(value = "基本单位", index = 36)
    private String entity;

    @ExcelProperty(value = "生产日期", index = 37)
    private String producitionDate;

    @ExcelProperty(value = "有效期", index = 38)
    private String PD;

    @ExcelProperty(value = "打包配置", index = 39)
    private String packageConfiguration;
}
