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

 * @Description:excel��������

 * @author:Administrator

 * @time:2017��8��21�� ����3:28:42

 */
public class ExcelUtil {
	
	/**
	
	 * @Description:Excel����
	
	 * @throws IOException
	
	 * void
	
	 * @exception:
	
	 * @author: Administrator
	
	 * @time:2017��8��21�� ����4:06:27
	
	 */
	public static void exportExcel(String xlsPath) throws IOException{
		
		xlsPath="E:\\test1\\12.xlsx";
		
		Integer index=xlsPath.indexOf(".");
		
		//�ж�����ļ���ʲô��ʽ
		String name=xlsPath.substring(index+1,xlsPath.length());
		
		Workbook workbook=null;
		
		if (name.equals("xlsx")) {
			workbook=new XSSFWorkbook();
		}else{
			workbook= new HSSFWorkbook();
		}
        //�½�������
        Sheet sheet = workbook.createSheet("hello");
        
        List<User> list=new ArrayList<User>();
        User user1=new User();
        user1.setId(10);
        user1.setName("��˧1");
        user1.setAge(21);
        
        User user2=new User();
        user2.setId(11);
        user2.setName("��˧2");
        user2.setAge(22);
        
        User user3=new User();
        user3.setId(12);
        user3.setName("��˧3");
        user3.setAge(23);
        
        list.add(user1);
        list.add(user2);
        list.add(user3);
        
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("id");
        row.createCell(1).setCellValue("����");
        row.createCell(2).setCellValue("����");
        //������,�к���Ϊ�������ݸ�createRow()����,��һ�д�0��ʼ����
        for (int i = 0; i < list.size(); i++) {
        	  row = sheet.createRow(i+1);
        	 for(int k = 0; k <list.size(); k++){
        		 row.createCell(0).setCellValue(list.get(i).getId().toString());
                 row.createCell(1).setCellValue(list.get(i).getName().toString());
                 row.createCell(2).setCellValue(list.get(i).getAge().toString());
        	 }
		}
        //������Ԫ��,row�Ѿ�ȷ�����к�,�к���Ϊ�������ݸ�createCell(),��һ�д�0��ʼ����
        FileOutputStream fos = new FileOutputStream(new File(xlsPath));
        workbook.write(fos);
        fos.close();
        
    }
	
	/**
	
	 * @Description:Excel����
	
	
	 * void
	
	 * @exception:
	
	 * @author: Administrator
	 * @throws IOException 
	 * @time:2017��8��21�� ����4:06:39
	
	 */
	public static List<User> readExcel(String xlsPath) throws IOException{
		xlsPath="E:\\test1\\12.xlsx";
		List<User> users=new ArrayList<User>();
		FileInputStream fileIn = new FileInputStream(xlsPath);  
		//����ָ�����ļ�����������Excel�Ӷ�����Workbook����  
		Workbook wb0 = new XSSFWorkbook(fileIn);  
		//��ȡExcel�ĵ��еĵ�һ����  
		Sheet sht0 = wb0.getSheetAt(0);  
		//��Sheet�е�ÿһ�н��е���  
		        for (Row r : sht0) {  
		        //�����ǰ�е��кţ���0��ʼ��δ�ﵽ2�������У������ѭ��  
		if(r.getRowNum()<1){  
		continue;  
		}  
		//����ʵ����  
		User user=new User();
		//ȡ����ǰ�е�1����Ԫ�����ݣ�����װ��infoʵ��stuName������  
		
		user.setId(Integer.valueOf(r.getCell(0).getStringCellValue()));
		user.setName(r.getCell(1).getStringCellValue());
		user.setAge(Integer.valueOf(r.getCell(2).getStringCellValue()));
		users.add(user);  
		        }  
		fileIn.close();      
		return users;
	}

}
