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
    public final static int exPayDaysDiff = 3;
    public final static String crmProductType = "外贸通";
    public final static String edmProductType="邮件营销";

    public static void main(String[] args) throws IOException {
        // 订单包含订单号，企业订单号，付款时间，产品类型，产品数量，购买月份，订单总价，订单客户名，订单客户序号，订单客户域名
        //客户包含：公司名、域名、企业序号、订单列表
        ArrayList<Custom> customsList = new ArrayList<>();
        //给到一个有数的表，能够输出一个客户列表
        FileInputStream sourceExcel = new FileInputStream("D:\\叶磊网易\\数据分析\\daima\\test\\order_name.xlsx");
        //导出增购订单表位置
        String outPath = "D:\\叶磊网易\\数据分析\\daima\\test\\exorder.xlsx";
        XSSFWorkbook workbook = new XSSFWorkbook(sourceExcel);
        XSSFSheet sheet = workbook.getSheetAt(0);
        //将Excel的订单数据转成客户对象列表
        extractExcelToCustomList(customsList, sheet);
        //制作订单对象列表
        ArrayList<Order> orderList = new ArrayList<>();
        exOrderList(customsList, orderList);
        //输出订单表格
        outputExOrderExcel(outPath, orderList);
        //关闭流
        sourceExcel.close();
    }

    /**
     * 将客户列表中的订单，梳理成订单对象列表
     *
     * @param customsList
     * @param orderList
     */
    private static void exOrderList(ArrayList<Custom> customsList, ArrayList<Order> orderList) {
        for (Custom ct1 : customsList) {
            ArrayList<Order> orderL1 = ct1.getOrderList();
            for (Order order : orderL1) {
                order.setCorpId(ct1.getCorpId());
                order.setCorpName(ct1.getCorpName());
                order.setDomain(ct1.getDomain());
                orderList.add(order);
            }
        }
    }

    private static void extractExcelToCustomList(ArrayList<Custom> customsList, XSSFSheet sheet) {
        ArrayList<String> corpNameList = new ArrayList<>();
        //每一行循环获取每行第一个cell的值（客户名称）
        for (int j = 1; j < sheet.getLastRowNum(); j++) {
            //获取行
            XSSFRow row = sheet.getRow(j);
            //获取行第一个单元格
            Cell cell0 = row.getCell(0);
            //将这个公司名录入公司名列表，并跳过重复值，录入到Custom列表
            String corpName = cell0.getStringCellValue();
            //将这一行的订单数据创建一个Order对象
            Order order = new Order(row.getCell(2).toString(),
                    row.getCell(4).toString(),
                    row.getCell(5).getDateCellValue(),
                    row.getCell(8).toString(),
                    Integer.parseInt(row.getCell(9).getStringCellValue()),
                    Integer.parseInt(row.getCell(10).getRawValue()),
                    row.getCell(12).getRawValue(),
                    row.getCell(1).getStringCellValue(), cell0.getStringCellValue(),
                    Integer.parseInt(row.getCell(3).getStringCellValue()));
            //把订单导入进客户列表内
            addOrderInCustom(customsList, corpNameList, row, corpName, order);
        }
    }

    /**
     * @param customs
     * @param corpNameList
     * @param row
     * @param corpName
     * @param order
     */
    private static void addOrderInCustom(ArrayList<Custom> customs,
                                         ArrayList<String> corpNameList,
                                         XSSFRow row, String corpName, Order order) {
        Date orderPaytime = order.getPaytime();
        //如果订单中的客户名，不在客户名单内
        if (!corpNameList.contains(corpName)) {
            // 那么就加入客户名单
            corpNameList.add(corpName);
            String domain = order.getDomain();
            int corpId = order.getCorpId();
            Date firstOrderDate = orderPaytime;
            // 并且生成一个客户添加到客户对象表// 添加该订单到客户对象表内的订单对象表
            customs.add(new Custom(corpName, domain, corpId, firstOrderDate, order));
        } else {
            //找到该客户，并且添加订单，需要给订单标记
            for (int i = 0; i < customs.size(); i++) {
                Custom custom = customs.get(i);
                //遍历找到该客户,公司名一致
                if (custom.getCorpName().equals(corpName)) {
                    //得到该客户custom，需要和之前的订单做比对
                    ArrayList<Order> orderList = custom.getOrderList();
                    //客户第一笔订单时间，新订单的付款时间
                    Date ctFirstOrderDate = custom.getFirstOrderDate();
                    Date orderPayDate = order.getPaytime();
                    //先比较这个订单的时间，和客户首次付款时间，在3天内的订单，就直接添加到订单列表，不需要标记
                    if (dateDiffDays(ctFirstOrderDate, orderPayDate) < exPayDaysDiff + 1) {
                        orderList.add(order);
                    } else {
                        //和客户首次付款时间超过3天，order则为增购，并且需要标记增购的类型
                        ordTypeTag(order, orderList, orderPayDate);
                        orderList.add(order);
                    }
                }
            }
        }
    }
    private static void ordTypeTag(Order order, ArrayList<Order> orderList, Date orderPayDate) {
        switch (order.getProductType()) {
            case crmProductType:
                for (Order oldOrder : orderList) {
                    if (dateDiffDays(oldOrder.getPaytime(), orderPayDate) == 0) {
                        //新订单和之前的所有老订单比对，如果发现和老订单是同一天的，则需要标记新老订单
                        //如果这个老订单是”外贸通“，则标记为exCRM
                        if (oldOrder.getProductType().equals(crmProductType)) {
                            order.setOrderType("exCRM");
                        } else if (oldOrder.getProductType().equals(edmProductType)) {
                            //如果是”邮件营销“，标记为exAll
                            order.setOrderType("exAll");
                            if (oldOrder.getOrderType() == null | oldOrder.getOrderType().equals("exEDM") | oldOrder.getOrderType().equals("exAll")) {
                                //如果邮件营销还是exEDM或是null或者是exall，则修改为exAll。
                                oldOrder.setOrderType("exAll");
                            }
                        }
                        //每个老订单都要对比一下
                        continue;
                    }
                    //没有找到订单列表内3天内的订单，就直接标记exCRM
                    order.setOrderType("exCRM");
                }
                break;
            //订单的类型是”邮件营销“
            case edmProductType:
                for (Order oldOrder : orderList) {
                    //如果找到超过3天的订单
                    if (dateDiffDays(oldOrder.getPaytime(), orderPayDate) == 0) {

                        if (oldOrder.getProductType().equals(edmProductType)) {
                            //如果这个老订单是”邮件营销“，则标记为exEDM
                            order.setOrderType("exEDM");
                        } else if (oldOrder.getProductType().equals(crmProductType)) {
                            //如果劳动是”外贸通“，标记为exAll
                            order.setOrderType("exAll");
                            if (oldOrder.getOrderType() == null | oldOrder.getOrderType().equals("exCRM")) {
                                // 如果外贸通该还是exCRM或是null，则修改为exAll。
                                oldOrder.setOrderType("exAll");
                            }
                        }
                        continue;
                    }
                    // 而且没有找到订单列表内3天内的订单，就直接标记exEDM
                    order.setOrderType("exEDM");
                }
                break;
            default:
                order.setOrderType("ex");
                break;
        }
    }

    private static int dateDiffDays(Date date01, Date date02) {
        return abs((int) ((date01.getTime() - date02.getTime()) / (1000 * 3600 * 24)));
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