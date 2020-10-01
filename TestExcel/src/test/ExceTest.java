package test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


//Excel書き込み

public class ExceTest {

	public static void main (String[]args) {


		XSSFWorkbook workbook1;
		try {

			//Excelファイルの場所
			workbook1= new XSSFWorkbook(new FileInputStream("C:\\Users\\onumaayano1199\\Pictures\\Sample1.xlsx"));

			//書き込みたいシート
			XSSFSheet sheet1=workbook1.getSheet("sheet1");

			//どこの行？
			XSSFRow row1=sheet1.createRow(1);

			//どこの列？
			XSSFCell cell1=row1.createCell(1);

			//書き込みたいこと
			cell1.setCellValue("aaa");

			//ファイルに返す
			FileOutputStream out1=new FileOutputStream("C:\\Users\\onumaayano1199\\Pictures\\Sample1.xlsx");

			//ファイルに書き込みます
			workbook1.write(out1);

		}catch(FileNotFoundException e) {
			e.printStackTrace();
		}catch(IOException e){
			e.printStackTrace();
		}
	}

}
