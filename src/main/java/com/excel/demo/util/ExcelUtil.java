package com.excel.demo.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.excel.demo.pojo.User;

/**

 * @Description:excel导出导入

 * @author:Administrator

 * @time:2017年8月21日 下午3:28:42

 */
public class ExcelUtil {
	
	public static void exportExcel() throws IOException{
		 //创建新工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //新建工作表
        HSSFSheet sheet = workbook.createSheet("hello");
        
        List<User> list=new ArrayList<User>();
        User user1=new User();
        user1.setId(10);
        user1.setName("陈帅1");
        user1.setAge(21);
        
        User user2=new User();
        user2.setId(11);
        user2.setName("陈帅2");
        user2.setAge(22);
        
        User user3=new User();
        user3.setId(12);
        user3.setName("陈帅3");
        user3.setAge(23);
        
        list.add(user1);
        list.add(user2);
        list.add(user3);
        
        HSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("id");
        row.createCell(1).setCellValue("姓名");
        row.createCell(2).setCellValue("年龄");
        //创建行,行号作为参数传递给createRow()方法,第一行从0开始计算
        for (int i = 0; i < list.size(); i++) {
        	  row = sheet.createRow(i+1);
        	 for(int k = 0; k <list.size(); k++){
        		 row.createCell(0).setCellValue(list.get(i).getId().toString());
                 row.createCell(1).setCellValue(list.get(i).getName().toString());
                 row.createCell(2).setCellValue(list.get(i).getAge().toString());
        	 }
		}
        //创建单元格,row已经确定了行号,列号作为参数传递给createCell(),第一列从0开始计算
        FileOutputStream fos = new FileOutputStream(new File("E:\\test1\\11.xls"));
        workbook.write(fos);
        fos.close();
    }

}
