package com.ghb.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelListener extends AnalysisEventListener<WriteModel> {
    private List<WriteModel> list = new ArrayList<WriteModel>();
    private List<WriteModel> list2 = new ArrayList<WriteModel>();

    private static String excelName;

    public void invoke(WriteModel data, AnalysisContext analysisContext) {
        list.add(data);
    }

    public void doAfterAllAnalysed(AnalysisContext analysisContext) {


        for (WriteModel writeModel : list) {
            if (writeModel.getSubCustCode() != null) {
                WriteModel w2 = new WriteModel();
                /**
                 * 入库时间
                 */
                String[] split = writeModel.getStorageTime().split("-");
                StringBuffer sb = new StringBuffer();
                w2.setStorageTime(sb.append(split[0]+"/"+split[1]+"/"+split[2]).toString());
                /**
                 * 主客户编码
                 */
                w2.setCustCode("");
                /**
                 * 主客户名称
                 */
                w2.setCustName(writeModel.getSubCustCode());
                excelName = writeModel.getSubCustCode();
                /**
                 * 分客户编码
                 */
                w2.setSubCustCode("");
                /**
                 * 分客户名称
                 */
                w2.setSubCustName(writeModel.getSubCustName());
                /**
                 * 进仓编号
                 */
                w2.setCustId(writeModel.getCustId());
                /**
                 * PO号
                 */
                w2.setPoNo(writeModel.getPoNo());
                /**
                 * SKU
                 */
                w2.setSkuNo(writeModel.getTypeNo());
                /**
                 * item
                 */
                w2.setItem(writeModel.getItem());
                /**
                 * 唛头
                 */
                w2.setMarkNo(writeModel.getLPNCode());
                /**
                 * 款号
                 */
                w2.setTypeNo(writeModel.getSkuNo());
                /**
                 * 色号
                 */
                w2.setColorNo(writeModel.getMarkNo());
                /**
                 * 托盘号
                 */
                w2.setLPNCode(writeModel.getCustType());
                /**
                 * 库位
                 */
                if (writeModel.getBatchNo() != null) {
                    String str = writeModel.getBatchNo();
                    if (str.charAt(3) == '0' & str.charAt(4) == '1') {
                        char[] chars = str.toCharArray();
                        char[] charss = new char[7];
                        charss[0] = chars[0];
                        charss[1] = chars[1];
                        charss[2] = chars[2];
                        charss[3] = chars[5];
                        charss[4] = chars[6];
                        charss[5] = '0';
                        charss[6] = chars[7];
                        String ss = String.valueOf(charss);
                        w2.setLoactionCode(ss);
                    } else {
                        w2.setLoactionCode(writeModel.getBatchNo());
                    }
                } else {
                    w2.setLoactionCode(writeModel.getBatchNo());
                }
                //处理长宽高
                if (writeModel.getQty_on_hand() != null) {

                    String[] str = writeModel.getQty_on_hand().split("\\*");

                    /**
                     * 长
                     */
                    w2.setLength(str[0]);
                    /**
                     * 宽
                     */
                    w2.setWidth(str[1]);
                    /**
                     * 高
                     */
                    w2.setHigh(str[2]);
                }
                /**
                 * 重量
                 */
                w2.setWeight("0.1");
                /**
                 * 现有数量
                 */
                w2.setQty_on_hand(writeModel.getColorNo());
                /**
                 * 尺码
                 */
                w2.setSizeNo("");
                /**
                 * 入库单号
                 */
                w2.setOrderNo(writeModel.getLength());
                /**
                 * 打包配置
                 */
                w2.setPackageConfiguration("1E0C0L");
                /**
                 * 商品名称
                 */
                w2.setCommodityName("货代标准产品");
                /**
                 * 有效期
                 */
                w2.setPD("");
                /**
                 * 商品品质
                 */
                w2.setQuality("正常");
                /**
                 * 库存状态
                 */
                w2.setStatus("可用");
                /**
                 * 商品编码
                 */
                w2.setProductCode("SP201906050001");
                /**
                 * 商品条码
                 */
                w2.setBarCode("hd0605001");
                /**
                 * 生产日期
                 */
                w2.setProducitionDate("");
                /**
                 * 供应商条码
                 */
                w2.setSupplierCode("");
                /**
                 * 供应商名称
                 */
                w2.setEntity("箱");
                w2.setSupplierName("");
                w2.setBarCode("hd0605001");
                w2.setCustProductCode("hd0605001");
                w2.setAttribute08("");
                w2.setAttribute09("");
                w2.setAttribute10("");
                w2.setAttribute11("");
                w2.setAttribute12("");
                list2.add(w2);
            }
        }
        String fileName = "C:\\Users\\ghb\\Desktop\\" + excelName  + new SimpleDateFormat("YYYY-MM-dd").format(new Date()) + "库存导入整理.xlsx";
        EasyExcel.write(fileName, WriteModel.class).sheet("模板").registerWriteHandler(new CustomCellWriteHandler()).doWrite(list2);
        System.out.println("库存导入模板生成完毕");
    }
}
