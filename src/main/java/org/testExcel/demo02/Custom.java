package org.testExcel.demo02;

import java.util.ArrayList;

public class Custom {
    String corpName;
    String domain;
    int corpId;
    ArrayList<Order> orderList = new ArrayList<>();

    public Custom(String corpName, String domain, int corpId, Order order) {
        this.corpName = corpName;
        this.domain = domain;
        this.corpId = corpId;
        this.orderList.add(order);
    }

    public Custom(String corpName, String domain, int corpId, ArrayList<Order> orderList) {
        this.corpName = corpName;
        this.domain = domain;
        this.corpId = corpId;
        this.orderList = orderList;
    }

    public Custom(String corpName, String domain, int corpId) {
        this.corpName = corpName;
        this.domain = domain;
        this.corpId = corpId;
    }

    public Custom(String corpName) {
        this.corpName = corpName;
    }

    public Custom() {
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

    public ArrayList<Order> getOrderList() {
        return orderList;
    }

    public void setOrderList(ArrayList<Order> orderList) {
        this.orderList = orderList;
    }

    @Override
    public String toString() {
        return "Custom{" +
                "corpName='" + corpName + '\'' +
                ", domain='" + domain + '\'' +
                ", corpId=" + corpId +
                ", orderList=" + orderList +
                '}';
    }
}
