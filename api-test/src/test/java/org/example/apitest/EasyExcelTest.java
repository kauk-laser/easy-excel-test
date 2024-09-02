package org.example.apitest;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.listener.ReadListener;
import org.junit.Assert;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;

import java.io.IOException;
import java.io.InputStream;


public class EasyExcelTest {
    private String directoryPath = "D:\\idea_workbench\\easy-excel-test\\api-test\\src\\test\\resources";

    private void readExcel(String resourcePath, Class clazz, ReadListener listener,String sheetName,Integer sheetNo){
        InputStream resourceAsStream = getClass().getClassLoader().getResourceAsStream(resourcePath);
        if (resourceAsStream != null){
            EasyExcel.read(resourceAsStream, clazz, listener).sheet(sheetNo,sheetName).doRead();
        }else {
            System.out.println("要读取的文件不存在");
        }
    }

    private void readExcel(String resourcePath, Class clazz, ReadListener listener){
       readExcel(resourcePath, clazz, listener,null,null);
    }

    @Test
    public void read_readBlankExcel_returnBlankString() {
        DemoDataListener demoDataListener = new DemoDataListener();
        readExcel("excel/blank.xlsx", String.class, demoDataListener);

        Assert.assertEquals(0, demoDataListener.getReadCount());
    }

    @Test
    public void read_readOneRowAndOneColoumExcel_returnTure() throws IOException {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/RowsAndColoums.xlsx", DemoData.class, demoDataListener);

        Assert.assertEquals(8, demoDataListener.getReadCount());
    }


}
