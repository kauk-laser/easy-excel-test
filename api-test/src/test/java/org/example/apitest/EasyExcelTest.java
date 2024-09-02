package org.example.apitest;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.listener.ReadListener;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.InputStream;


public class EasyExcelTest {
    private  <T>  void readExcel(String resourcePath, Class<T> clazz, ReadListener<T> listener, String sheetName, Integer sheetNo){
        InputStream resourceAsStream = getClass().getClassLoader().getResourceAsStream(resourcePath);
        if (resourceAsStream != null){
            EasyExcel.read(resourceAsStream, clazz, listener).sheet(sheetNo,sheetName).doRead();
        }else {
            System.out.println("要读取的文件不存在");
        }
    }

    private <T> void readExcel(String resourcePath, Class<T> clazz, ReadListener<T> listener){
       readExcel(resourcePath, clazz, listener,null,null);
    }

    @Test
    public void read_readBlankExcel_returnBlank() {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/blank.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(0, demoDataListener.getReadCount());
    }

    @Test
    public void read_readOnlyColoumNameExcel_returnZero() {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/onlyColoumName.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(0, demoDataListener.getReadCount());
    }

    @Test
    public void read_readOnlyOneRowDataExcel_returnOne() {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/oneRowData.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(1, demoDataListener.getReadCount());
    }

    @Test
    public void read_readNotAdjacentColoums_returnBlankString() {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/notAdjacentColoums.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(1, demoDataListener.getReadCount());
    }

    @Test
    public void read_readRowsAndColoumsExcel_returnTure() {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/RowsAndColoums.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(8, demoDataListener.getReadCount());
    }


}
