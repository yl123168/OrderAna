package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class Main {
    public static void main(String[] args) throws IOException {
        String path = "D:\\叶磊网易\\数据分析\\daima\\yuchao\\emaillist.xlsx";
        String outFile = "D:\\叶磊网易\\数据分析\\daima\\yuchao\\emailListOut.xlsx";
        FileInputStream fileInputStream = new FileInputStream(path);
        //获取默认的sheet
        XSSFWorkbook sourceWb = new XSSFWorkbook(fileInputStream);
        XSSFSheet activeSheet = sourceWb.getSheetAt(0);
        ArrayList<String> domainList = new ArrayList<>();
        for (int j = 1; j < activeSheet.getLastRowNum(); j++) {
            XSSFRow row = activeSheet.getRow(j);
            Cell cellName = row.getCell(4);
            String s = cellName.toString();
            if(!domainList.contains(s)){
                domainList.add(s);
            }
        }
        //得到一个
        ArrayList<String> domainName = new ArrayList<>();
        //按照

        fileInputStream.close();
    }
}