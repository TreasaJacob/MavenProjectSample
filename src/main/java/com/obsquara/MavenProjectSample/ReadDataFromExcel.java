package com.obsquara.MavenProjectSample;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ReadDataFromExcel {
	
		static FileInputStream fi;
		static XSSFWorkbook wb;
		static XSSFSheet sh;
		
		public static String readName (int i,int j) throws Exception
		{
			fi = new FileInputStream("C://Assignments//ReadDataEg.xlsx");
			wb = new XSSFWorkbook(fi);
			sh = wb.getSheet("Data Sheet");
			XSSFRow row = sh.getRow(i);
			XSSFCell cell = row.getCell(j);
			return cell.getStringCellValue();
			
			
		}
		public static double readRollNo( int i,int j) throws Exception
		{
			
			fi = new FileInputStream("C://Assignments//ReadDataEg.xlsx");
			wb = new XSSFWorkbook(fi);
			sh = wb.getSheet("Data Sheet");
			XSSFRow row = sh.getRow(i);
			XSSFCell cell = row.getCell(j);
			return cell.getNumericCellValue();
		}
		public static String readClass(int i, int j) throws Exception
		{
			fi = new FileInputStream("C://Assignments//ReadDataEg.xlsx");
			wb = new XSSFWorkbook(fi);
			sh = wb.getSheet("Data Sheet");
			XSSFRow row = sh.getRow(i);
			XSSFCell cell = row.getCell(j);
			return cell.getStringCellValue();
			
		}
	

	public static void main(String[] args) throws Exception {
		String value1 = ReadDataFromExcel.readName(1,0);
		System.out.println(value1);
		double  value2 = ReadDataFromExcel.readRollNo(1, 1);
		System.out.println(value2);
		String value3 = ReadDataFromExcel.readClass(1, 2);
		System.out.println(value3);


	}

}
