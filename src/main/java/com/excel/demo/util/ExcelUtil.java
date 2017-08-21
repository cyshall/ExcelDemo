package com.excel.demo.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.Thread.State;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.excel.demo.pojo.User;

/**

 * @Description:excel导出导入

 * @author:Administrator

 * @time:2017年8月21日 下午3:28:42

 */
public class ExcelUtil {
	
	/**
	
	 * @Description:Excel导出
	
	 * @throws IOException
	
	 * void
	
	 * @exception:
	
	 * @author: Administrator
	
	 * @time:2017年8月21日 下午4:06:27
	
	 */
	public static void exportExcel(String xlsPath) throws IOException{
		
		xlsPath="E:\\test1\\12.xlsx";
		
		Integer index=xlsPath.indexOf(".");
		
		//判断这个文件是什么格式
		String name=xlsPath.substring(index+1,xlsPath.length());
		
		Workbook workbook=null;
		
		if (name.equals("xlsx")) {
			workbook=new XSSFWorkbook();
		}else{
			workbook= new HSSFWorkbook();
		}
        //新建工作表
        Sheet sheet = workbook.createSheet("hello");
        
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
        
        Row row = sheet.createRow(0);
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
        FileOutputStream fos = new FileOutputStream(new File(xlsPath));
        workbook.write(fos);
        fos.close();
        
    }
	
	/**
	
	 * @Description:Excel导入
	
	
	 * void
	
	 * @exception:
	
	 * @author: Administrator
	 * @throws IOException 
	 * @time:2017年8月21日 下午4:06:39
	
	 */
	public static List<User> readExcel(String xlsPath) throws IOException{
		xlsPath="E:\\test1\\12.xlsx";
		List<User> users=new ArrayList<User>();
		FileInputStream fileIn = new FileInputStream(xlsPath);  
		//根据指定的文件输入流导入Excel从而产生Workbook对象  
		Workbook wb0 = new XSSFWorkbook(fileIn);  
		//获取Excel文档中的第一个表单  
		Sheet sht0 = wb0.getSheetAt(0);  
		//对Sheet中的每一行进行迭代  
		        for (Row r : sht0) {  
		        //如果当前行的行号（从0开始）未达到2（第三行）则从新循环  
		if(r.getRowNum()<1){  
		continue;  
		}  
		//创建实体类  
		User user=new User();
		//取出当前行第1个单元格数据，并封装在info实体stuName属性上  
		
		user.setId(Integer.valueOf(r.getCell(0).getStringCellValue()));
		user.setName(r.getCell(1).getStringCellValue());
		user.setAge(Integer.valueOf(r.getCell(2).getStringCellValue()));
		users.add(user);  
		        }  
		fileIn.close();      
		return users;
	}

}
