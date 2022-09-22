package org.testExcel.demo02;

import java.util.Date;
import java.util.Objects;

public class Order {
    String orderId;
    String orderCorpId;
    Date paytime;
    String productType;
    int skuNumber;
    int skuMonth;
    String orderPrice;
    String orderType;
    String corpName;
    String domain;
    int corpId;
    Date firstOrderDate;

    public Date getFirstOrderDate() {
        return firstOrderDate;
    }

    public void setFirstOrderDate(Date firstOrderDate) {
        this.firstOrderDate = firstOrderDate;
    }

    public String getOrderType() {
        return orderType;
    }

    public void setOrderType(String orderType) {
        this.orderType = orderType;
    }

    public String getOrderId() {
        return orderId;
    }

    public void setOrderId(String orderId) {
        this.orderId = orderId;
    }

    public String getOrderCorpId() {
        return orderCorpId;
    }

    public void setOrderCorpId(String orderCorpId) {
        this.orderCorpId = orderCorpId;
    }

    public Date getPaytime() {
        return paytime;
    }

    public void setPaytime(Date paytime) {
        this.paytime = paytime;
    }

    public String getProductType() {
        return productType;
    }

    public void setProductType(String productType) {
        this.productType = productType;
    }

    public int getSkuNumber() {
        return skuNumber;
    }

    public void setSkuNumber(int skuNumber) {
        this.skuNumber = skuNumber;
    }

    public int getSkuMonth() {
        return skuMonth;
    }

    public void setSkuMonth(int skuMonth) {
        this.skuMonth = skuMonth;
    }

    public String getCorpName() {
        return corpName;
    }

    public void setCorpName(String corpName) {
        this.corpName = corpName;
    }

    public String getDomain() {
        return domain;
    }

    public void setDomain(String domain) {
        this.domain = domain;
    }

    public int getCorpId() {
        return corpId;
    }

    public void setCorpId(int corpId) {
        this.corpId = corpId;
    }

    public String getOrderPrice() {
        return orderPrice;
    }

    public void setOrderPrice(String orderPrice) {
        this.orderPrice = orderPrice;
    }

    public Order(String orderId, String orderCorpId, Date paytime, String productType,
                 int skuNumber, int skuMonth, String orderPrice, String domain,String corpName,int corpId) {
        this.orderId = orderId;
        this.orderCorpId = orderCorpId;
        this.paytime = paytime;
        this.productType = productType;
        this.skuNumber = skuNumber;
        this.skuMonth = skuMonth;
        this.orderPrice = orderPrice;
        this.domain = domain;
        this.corpName = corpName;
        this.corpId = corpId;
    }
}
