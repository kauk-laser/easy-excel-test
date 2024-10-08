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
        /*
          读取失败
          easyExcel默认sheet中的第一行必须是表头/列名
          所以这里会抛出异常
          因为它尝试把sheet中的不在第一行的列名作为数据行读取并转换，发生类型转换错误
         */
        DemoDataListener demoDataListener = new DemoDataListener();
        Assert.assertThrows(ExcelDataConvertException.class,()-> readExcel("excel/coloumNameNotAtFirstRow.xlsx", DemoData.class, demoDataListener));
    }

    @Test
    public void read_designateRowOfColumnName_ReadSuccess() {
        /*
          可以使用sheet的headRowNumber方法指定读取的列名所在的行，来支持第一行不是表头的情况
         */
        DemoDataListener demoDataListener = new DemoDataListener();
        InputStream resourceAsStream = getClass().getClassLoader().getResourceAsStream("excel/coloumNameNotAtFirstRow.xlsx");
        if (resourceAsStream != null){
            EasyExcel.read(resourceAsStream, DemoData.class, demoDataListener).sheet().headRowNumber(5).doRead();
        }else {
            System.out.println("要读取的文件不存在");
        }
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
    public void read_designatedBySheetNo_returnDataCountInDesignatedSheet() {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/designatedBySheetNo.xlsx", DemoData.class, demoDataListener,null,1);

        Assertions.assertEquals(1, demoDataListener.getReadCount());
    }

    @Test
    public void read_noColumnNames_returnNoData() {
        /*
          读取失败
          因为没有列名，所以无法读取数据
         */
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/noColumnNames.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(0, demoDataListener.getReadCount());
    }

    @Test
    public void read_RowsAndColoums_returnTure() {
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/RowsAndColoums.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(8, demoDataListener.getReadCount());
    }

    @Test
    public void read_OneOfColumnMissed_readSuccessAndUsingDefaultValueOnMissedColumn(){
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/OneOfColumnMissed.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(8, demoDataListener.getReadCount());
    }

    @Test
    public void read_twoTablesAndHaveOneColumnNameIsSame_ReadOneTableAtOnceAndAlawaysReadColumnOccursFirst() {
        /*
          读取成功，easyExcel通过DemoData包含的字段来决定需要读取的数据
          但是要注意，如果有两个相同的列名，那么只会读取第一个
         */
        DemoDataListener demoDataListener = new DemoDataListener();

        readExcel("excel/twoTables.xlsx", DemoData.class, demoDataListener);

        Assertions.assertEquals(8, demoDataListener.getReadCount());

        StudentDataListener studentDataListener = new StudentDataListener();

        readExcel("excel/twoTables.xlsx", StudentData.class, studentDataListener);

        Assertions.assertEquals(8, demoDataListener.getReadCount());
    }


}
