package Exceltest;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

//Excel読み込み

public class Testinput {

	public static void main(String[]args) {

		//変数初期化
		InputStream is =null;
		Workbook wb=null;

		try {

			//読み込みたいファイル
			is=new FileInputStream("C:\\Users\\onumaayano1199\\Pictures\\Sample1.xlsx");
			wb=WorkbookFactory.create(is);

			//どこのシート？
			Sheet sh =wb.getSheet("testsheet2");

			//どこの行？
			Row row =sh.getRow(0);

			//どこの列？
			Cell cell=row.getCell(0);

			//指定の値取ってきます
			String value=cell.getStringCellValue();

			//別の値も取ってみます

			//どこの行？
			Row row1 =sh.getRow(1);

			//どこの列？
			Cell cell1=row1.getCell(1);

			//指定の値取ってきます
			String value1=cell1.getStringCellValue();
			
			
			//日付型の取得
			Row row2 =sh.getRow(4);
			Cell cell2 =row2.getCell(0);
			int type=cell2.getCellType();
			
			if(type ==Cell.CELL_TYPE_NUM) {
				
			}
			

			//コンソールに出力したい値を書く。
			System.out.println(value);
			System.out.println(value1);

		}catch(Exception ex) {
			ex.printStackTrace();
		}finally {
			try {
				wb.close();
			}catch(Exception ex2) {
				ex2.printStackTrace();
			}
		}
	}

}
