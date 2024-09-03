package org.example.apitest;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.exception.ExcelDataConvertException;
import com.alibaba.excel.read.listener.ReadListener;
import org.junit.Assert;
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
    public void read_OnlyColoumName_returnZero() {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/onlyColoumName.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(0, demoDataListener.getReadCount());
    }

    @Test
    public void read_OnlyOneRowData_returnOne() {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/oneRowData.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(1, demoDataListener.getReadCount());
    }

    @Test
    public void read_NotAdjacentColoums_readAllColoums() {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/notAdjacentColoums.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(1, demoDataListener.getReadCount());
    }

    @Test
    public void read_NotAdjacentRows_readAllRows() {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/notAdjacentRows.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(2, demoDataListener.getReadCount());
    }

    @Test
    public void read_coloumNameNotAtFirstRow_ReadFaild() {
        /**
         * 读取失败
         * easyExcel规定sheet中的第一行必须是表头/列名
         * 所以这里会抛出异常
         * 因为它尝试把sheet中的不在第一行的列名作为数据行读取并转换，发生类型转换错误
         */
        DemoDataListener demoDataListener = new DemoDataListener();
        Assert.assertThrows(ExcelDataConvertException.class,()->{
            readExcel("excel/coloumNameNotAtFirstRow.xlsx", DemoData.class, demoDataListener);
        });
    }

    @Test
    public void read_columnNameNotStartFromFirstColumn_ReadSuccess() {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/columnNameNotStartFromFirstColumn.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(1, demoDataListener.getReadCount());
    }

    @Test
    public void read_designatedBySheetName_returnDataCountInDesignatedSheet() {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/designatedBySheetName.xlsx", DemoData.class, demoDataListener,"Sheet2",null);

        Assertions.assertEquals(1, demoDataListener.getReadCount());
    }

    @Test
    public void read_RowsAndColoums_returnTure() {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/RowsAndColoums.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(8, demoDataListener.getReadCount());
    }


}
