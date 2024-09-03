package org.example.apitest;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
@EqualsAndHashCode
public class StudentData {
    @ExcelProperty("姓名")
    private String name;
    @ExcelProperty("班级")
    private int classNum;
    @ExcelProperty("分数")
    private int score;
}
