package org.example.apitest;

public class DemoDAO {
    public void save(Object object) {
        // 如果是mybatis,尽量别直接调用多次insert,自己写一个mapper里面新增一个方法batchInsert
        // Mybatis批量保存数据
    }
}
