package test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


//既存のExcelファイルの編集
public class Excelwrite {

	public static void main(String[] args) {

		//変数の初期化
		FileInputStream in=null;
		Workbook wb=null;

		try {
			//編集したいファイルの場所と名前
			in=new FileInputStream("C:\\Users\\onumaayano1199\\Pictures\\Sample1.xlsx");

			//編集機能へファイルを葬る
			wb =WorkbookFactory.create(in);

		}catch(IOException e) {
			System.out.println(e.toString());
		}finally {
			try {
				in.close();
			}catch(IOException e) {
				System.out.println(e.toString());
			}
		}

		//ファイル内に追加で新しいシートを作りたいとき
		//Sheet sheet=wb.createSheet("new sheet");

		//値を追加、更新したいときはこれ↓

		//①書き込みたいシート
		//Sheet sheet1=wb.getSheet("new sheet");
		Sheet sheet1=wb.getSheet("new sheet");

		//②どこの行？※1列目=0
		//Row row1=sheet1.createRow(1);
		Row row1=sheet1.createRow(1);

		//③どこの列？
		//Cell cell1=row1.createCell(1);
		Cell cell1=row1.createCell(2);

		//④書き込みたいこと
		//cell1.setCellValue("bbb");
		double time=DateUtil.convertTime("3:30");
		cell1.setCellValue(time);




		//やりたいことが終わったらファイルへ保存するよ～
		//変数の初期化
		FileOutputStream out =null;

		try {
			//ここに返します
			out=new FileOutputStream("C:\\Users\\onumaayano1199\\Pictures\\Sample1.xlsx");

			//編集部分を書いて保存しまーす
			wb.write(out);
		}catch(IOException e) {
			System.out.println(e.toString());
		}finally {
			try {
				out.close();
			}catch(IOException e) {
				System.out.println(e.toString());
			}
		}
	}

}
