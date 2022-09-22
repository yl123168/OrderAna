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
    public final static String edmProductType = "邮件营销";
    public final static String crmExTag = "exCRM";
    public final static String edmExTag = "exEDM";
    public final static String allExTag = "exAll";

    public static void main(String[] args) throws IOException {
        //订单包含订单号，企业订单号，付款时间，产品类型，产品数量，购买月份，订单总价，订单客户名，订单客户序号，订单客户域名
        //客户包含：公司名、域名、企业序号、订单列表
        ArrayList<Custom> customsList = new ArrayList<>();
//        System.out.println("输入订单列表源文件路径");
//        Scanner sc = new Scanner(System.in);
        String sourceExlPath = "D:\\叶磊网易\\数据分析\\订单数据\\增购订单分析\\orderList.xlsx";
        //导出增购订单表位置
        String outExlPath = "D:\\叶磊网易\\数据分析\\订单数据\\增购订单分析\\exOrder.xlsx";
        String outCustomExlPath = "D:\\叶磊网易\\数据分析\\订单数据\\增购订单分析\\customList.xlsx";
        //给到一个有数的表，能够输出一个客户列表
        FileInputStream sourceExcel = new FileInputStream(sourceExlPath);
        XSSFWorkbook workbook = new XSSFWorkbook(sourceExcel);
        XSSFSheet sheet = workbook.getSheetAt(0);
        //将Excel的订单数据转成客户对象列表
        extractExcelToCustomList(customsList, sheet);
        //制作订单对象列表
        ArrayList<Order> orderList = new ArrayList<>();
        exOrderList(customsList, orderList);
        //输出订单表格
        outputExOrderExcel(outExlPath, orderList);
        //输出客户台账(客户域名、公司名、序号、首次购买时间、购买外贸通数量，购买邮件营销数量)
        outputExCustomExcel(customsList, outCustomExlPath);
        //关闭流
        sourceExcel.close();
    }

    private static void outputExCustomExcel(ArrayList<Custom> customsList, String outPath) throws IOException {
        customListInit(customsList);
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        for (int i = 0; i < customsList.size(); i++) {
            XSSFRow row = sheet.createRow(i);
            Custom custom = customsList.get(i);
            row.createCell(0).setCellValue(custom.getDomain());
            row.createCell(1).setCellValue(custom.getCorpName());
            row.createCell(2).setCellValue(custom.getCorpId());
            row.createCell(3).setCellValue(custom.getFirstOrderDate());
            row.createCell(4).setCellValue(custom.getCrmAcountNum());
            row.createCell(5).setCellValue(custom.getEdmNum());
        }
        FileOutputStream fileOutputStream = new FileOutputStream(outPath);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    private static void customListInit(ArrayList<Custom> customList) {
        for (Custom custom : customList) {
            ArrayList<Order> orderList = custom.getOrderList();
            for (Order order : orderList) {
                if (order.getProductType().equals(crmProductType)) {
                    custom.setCrmAcountNum(custom.getCrmAcountNum() + order.getSkuNumber());
                } else if (order.getProductType().equals(edmProductType)) {
                    custom.setEdmNum(custom.getEdmNum() + order.getSkuNumber());
                }
            }
        }
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
        XSSFRow row1 = sheet.getRow(0);
        int indexOfcorpID = 0, indexOfcorpName = 0, indexOfDomainName = 0, indexOfOdCorpNumber = 0, indexOfOdNumber = 0,
                indexOfOdTime = 0, indexOfOdProductType = 0, indexOfSkuNum = 0, indexOfSkumouth = 0, indexOfOdPrice = 0;
        for (int i = 0; i < row1.getLastCellNum(); i++) {
            String value = row1.getCell(i).getStringCellValue();
            switch (value) {
                case "企业ID":
                    indexOfcorpID = i;
                    break;
                case "客户名称":
                    indexOfcorpName = i;
                    break;
                case "域名":
                    indexOfDomainName = i;
                    break;
                case "企业订单号":
                    indexOfOdCorpNumber = i;
                    break;
                case "订单号":
                    indexOfOdNumber = i;
                    break;
                case "付款时间":
                    indexOfOdTime = i;
                    break;
                case "产品类型":
                    indexOfOdProductType = i;
                    break;
                case "购买数量":
                    indexOfSkuNum = i;
                    break;
                case "平均值(month_buy)":
                    indexOfSkumouth = i;
                    break;
                case "订单总价":
                    indexOfOdPrice = i;
                    break;
            }
        }
        //每一行循环获取每行第一个cell的值（客户名称）
        for (int j = 1; j < sheet.getLastRowNum() + 1; j++) {
            //获取行
            XSSFRow row = sheet.getRow(j);
            //获取公司名所在单元格
            Cell cell0 = row.getCell(indexOfcorpName);
            //将这个公司名录入公司名列表，并跳过重复值，录入到Custom列表
            String corpName = cell0.getStringCellValue();

            //将这一行的订单数据创建一个Order对象
            Order order = new Order(row.getCell(indexOfOdNumber).getStringCellValue(),
                    row.getCell(indexOfOdCorpNumber).getStringCellValue(),
                    row.getCell(indexOfOdTime).getDateCellValue(),
                    row.getCell(indexOfOdProductType).getStringCellValue(),
                    Integer.parseInt(row.getCell(indexOfSkuNum).getStringCellValue()),
                    Integer.parseInt(row.getCell(indexOfSkumouth).getRawValue()),
                    row.getCell(indexOfOdPrice).getRawValue(),
                    row.getCell(indexOfDomainName).getStringCellValue(),
                    row.getCell(indexOfcorpName).getStringCellValue(),
                    Integer.parseInt(row.getCell(indexOfcorpID).getStringCellValue()));
            ;
            //把订单导入进客户列表内
            addOrderInCustom(customsList, order);
        }
    }

    /**判断是不是新客户，新的订单内包含的信息，和之前的客户是不是同一个
     *
     * @return
     */
    private static boolean isNewCustom(String corpName, int corpId, String domain, ArrayList<Custom> customs) {
        //遍历之前的客户对象表，如果是其中一个客户的信息，那么就是老客户，否则就是新客户
        for (Custom custom:customs) {
            if(isSameCustom(corpName,corpId, domain,custom)){
                return false;
            }
        }
        return true;
    }

    /**判断客户是不是一个，不同的企业ID，不同的公司名，不同的域名
     *
     * @param corpName
     * @param corpId
     * @param domain
     * @param custom
     * @return
     */
    private static boolean isSameCustom(String corpName, int corpId, String domain, Custom custom) {
        if(custom.getDomain().equals(domain) |custom.getCorpName().equals(corpName) | custom.getCorpId() == corpId){
            return true;
        }else {
            return false;
        }
    }

    /**把订单对象，添加到Custom对象列表，如果是新客户，直接添加一个新的客户对象，如果订单是老客户，把订单对象添加到老客户订单列表
     * @param customs
     * @param order
     */
    private static void addOrderInCustom(ArrayList<Custom> customs, Order order) {
        Date orderPaytime = order.getPaytime();

        String corpName = order.getCorpName();
        int corpId = order.getCorpId();
        String domain = order.getDomain();

        //新客户加入名单，老客户要把订单加入老客户内的订单表
        if (isNewCustom(corpName, corpId,domain,customs)) {
            Date firstOrderDate = orderPaytime;
            // 并且生成一个客户添加到客户对象表// 添加该订单到客户对象表内的订单对象表
            customs.add(new Custom(corpName, domain, corpId, firstOrderDate, order));
        } else {
            for (int i = 0; i < customs.size(); i++) {
                Custom custom = customs.get(i);
                //找到同一个客户,公司名一致
                if (isSameCustom(corpName,corpId,domain,custom)) {
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
        /*if (!corpNameList.contains(corpName)) {
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
            }*/
        }
    }

    /**订单增购类型判断，是单独增购外贸通、邮件营销，还是增购B+C
     *
     * @param order
     * @param orderList
     * @param orderPayDate
     */
    private static void ordTypeTag(Order order, ArrayList<Order> orderList, Date orderPayDate) {
        switch (order.getProductType()) {
            case crmProductType:
                for (Order oldOrder : orderList) {
                    if (dateDiffDays(oldOrder.getPaytime(), orderPayDate) == 0) {
                        //新订单和之前的所有老订单比对，如果发现和老订单是同一天的，则需要标记新老订单
                        //如果同一天老订单是”外贸通“，则标记为exCRM
                        if (oldOrder.getOrderType().equals(allExTag)) {
                            order.setOrderType(allExTag);
                        } else if (oldOrder.getProductType().equals(crmProductType) & oldOrder.getOrderType().equals(crmExTag)) {
                            order.setOrderType(crmExTag);
                        } else if (oldOrder.getProductType().equals(edmProductType)) {
                            //如果同一天的老订单是”邮件营销“，标记为exAll
                            order.setOrderType(allExTag);
                            oldOrder.setOrderType(allExTag);
                        }
                        //每个老订单都要对比一下
                        continue;
                    }
                    //没有找到订单列表内3天内的订单，就直接标记exCRM
                    order.setOrderType(crmExTag);
                }
                break;
            //订单的类型是”邮件营销“
            case edmProductType:
                for (Order oldOrder : orderList) {
                    //如果找到超过3天的订单
                    if (dateDiffDays(oldOrder.getPaytime(), orderPayDate) == 0) {
                        if (oldOrder.getOrderType().equals(allExTag)) {
                            order.setOrderType(allExTag);
                        } else if (oldOrder.getProductType().equals(edmProductType)) {
                            //如果这个老订单是”邮件营销“，则标记为exEDM
                            order.setOrderType(edmExTag);
                        } else if (oldOrder.getProductType().equals(crmProductType)) {
                            //如果老订单是”外贸通“，标记为exAll
                            order.setOrderType(allExTag);
                            oldOrder.setOrderType(allExTag);
                        }
                        continue;
                    }
                    // 而且没有找到订单列表内3天内的订单，就直接标记exEDM
                    order.setOrderType(edmExTag);
                }
                break;
            default:
                order.setOrderType("exError");
                break;
        }
    }

    /**
     * 获得两个日期之间相差的天数，取整数天
     *
     * @param date01
     * @param date02
     * @return
     */
    private static int dateDiffDays(Date date01, Date date02) {
        return abs((int) ((date01.getTime() - date02.getTime()) / (1000 * 3600 * 24)));
    }

    /**
     * 判断订单是否是增购的订单类型
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
}