package org.testExcel.demo02;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;

import static java.lang.Math.abs;

public class Main {
    public static void main(String[] args) throws IOException {
        // 订单包含订单号，企业订单号，付款时间，产品类型，产品数量，购买月份，订单总价，订单客户名，订单客户序号，订单客户域名
        //客户包含：公司名、域名、企业序号、订单列表
        ArrayList<Custom> customs = new ArrayList<>();
        //给到一个有数的表，能够输出一个客户列表
        FileInputStream sourceExcel = new FileInputStream("D:\\叶磊网易\\数据分析\\daima\\test\\order_name.xlsx");
        //导出增购订单表位置
        String outPath = "D:\\叶磊网易\\数据分析\\daima\\test\\exorder.xlsx";
        XSSFWorkbook workbook = new XSSFWorkbook(sourceExcel);
        XSSFSheet sheet = workbook.getSheetAt(0);
        //将Excel的订单数据转成客户对象列表
        extractExcelToCustomList(customs, sheet);

        //制作订单对象列表
        ArrayList<Order> orderList = new ArrayList<>();
        outputOrderList(customs, orderList);
        //输出订单表格
        outputExOrderExcel(outPath, orderList);
        //关闭流
        sourceExcel.close();
    }

    /**
     * 将客户列表中的订单，梳理成订单对象列表
     *
     * @param customs
     * @param orderList
     */
    private static void outputOrderList(ArrayList<Custom> customs, ArrayList<Order> orderList) {
        for (Custom ct1 : customs) {
            ArrayList<Order> orderL1 = ct1.getOrderList();
            for (Order order : orderL1) {
                order.setCorpId(ct1.getCorpId());
                order.setCorpName(ct1.getCorpName());
                order.setDomain(ct1.getDomain());
                orderList.add(order);
            }
        }
    }

    private static void extractExcelToCustomList(ArrayList<Custom> customs, XSSFSheet sheet) {
        ArrayList<String> corpNameList = new ArrayList<>();
        //每一行循环获取每行第一个cell的值（客户名称）
        for (int j = 1; j < sheet.getLastRowNum(); j++) {
            //获取行
            XSSFRow row = sheet.getRow(j);
            //获取行第一个单元格
            Cell cell0 = row.getCell(0);
            //将这个公司名录入公司名列表，并跳过重复值，录入到Custom列表
            String corpName = cell0.toString();
            Order order = new Order(row.getCell(2).toString(), row.getCell(4).toString(), row.getCell(5).getDateCellValue(), row.getCell(8).toString(),
                    (int) row.getCell(9).getNumericCellValue(), Integer.parseInt(row.getCell(10).getRawValue()), (row.getCell(12).getRawValue()));

            //如果这个订单的客户不在客户表，则添加客户并且添加订单
            if (!corpNameList.contains(corpName)) {
                corpNameList.add(corpName);
                String domain = row.getCell(1).toString();
                int corpId = (int) row.getCell(3).getNumericCellValue();
                customs.add(new Custom(corpName, domain, corpId, order));
            } else {
                //找到该客户，并且添加订单
                for (int i = 0; i < customs.size(); i++) {
                    Custom custom = customs.get(i);
                    //遍历找到了该订单
                    if (custom.getCorpName() == corpName) {
                        ArrayList<Order> orderList = custom.getOrderList();
                        for (Order order1 : orderList) {
                            //如果是不同的订单号并且间隔天数大于3天
                            if (isExtensionOrder(order, order1)) {
                                order.setOrderType(1);
                                //确认这个订单是增购订单
                            }
                            custom.getOrderList().add(order);
                            break;
                        }
                    }
                }
            }
        }
    }

    /**
     * 判断订单是否是增购的订单类型，如果是，则标记为1
     *
     * @param order1
     * @param order2
     * @return
     */
    private static boolean isExtensionOrder(Order order1, Order order2) {
        Date date2 = order1.getPaytime();
        Date date1 = order2.getPaytime();
        //如果是不同的订单号并且间隔天数大于3天
        if ((!order1.getOrderCorpId().equals(order2.getOrderCorpId())) &&
                abs((int) ((date1.getTime() - date2.getTime()) / (1000 * 3600 * 24))) > 3
        ) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * 将订单列表输出成一个Excel表格
     *
     * @param outPath
     * @param orderList
     * @throws IOException
     */
    private static void outputExOrderExcel(String outPath, ArrayList<Order> orderList) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        for (int i = 0; i < orderList.size(); i++) {
            XSSFRow row = sheet.createRow(i);
            Order order = orderList.get(i);
            row.createCell(0).setCellValue(order.getCorpName());
            row.createCell(1).setCellValue(order.getCorpId());
            row.createCell(2).setCellValue(order.getDomain());
            row.createCell(3).setCellValue(order.getPaytime());
            row.createCell(4).setCellValue(order.getOrderCorpId());
            row.createCell(5).setCellValue(order.getOrderId());
            row.createCell(6).setCellValue(order.getProductType());
            row.createCell(7).setCellValue(order.getSkuMonth());
            row.createCell(8).setCellValue(order.getSkuNumber());
            row.createCell(9).setCellValue(order.getOrderType());
            row.createCell(10).setCellValue(order.getOrderPrice());
        }
        FileOutputStream fileOutputStream = new FileOutputStream(outPath);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    //按照客户维度，每行客户关联其订单,暂时用不到
    /*private static void OutputExcel(String outPath, ArrayList<Custom> customs) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        for (int i = 0; i < customs.size(); i++) {
            XSSFRow row = sheet.createRow(i);
            Custom custom = customs.get(i);
            row.createCell(0).setCellValue(custom.getCorpName());
            row.createCell(1).setCellValue(custom.getCorpId());
            row.createCell(2).setCellValue(custom.getDomain());
            ArrayList<Order> orderList = custom.getOrderList();
            for (int j = 0; j < orderList.size(); j++) {
                Order order = orderList.get(j);
                row.createCell(3 + j * 4).setCellValue(order.getProductType());
                row.createCell(4 + j * 4).setCellValue(order.getPaytime());
                row.createCell(5 + j * 4).setCellValue(order.getOrderPrice());
                row.createCell(6 + j * 4).setCellValue(order.getOrderType());
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(outPath);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }*/
}