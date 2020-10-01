package Exceltest;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


//Excelファイル生成

public class WorkBook {
	public static void main(String[] args) {

	    //Excelデータ生成
		//xlsx形式のExcelを生成します
	    Workbook wb = new XSSFWorkbook();

	    //ファイル内にシートを作ります
	    Sheet sh = wb.createSheet();

	    //どこの行に書くの？
	    Row row = sh.createRow(0);

	    //どこの列に書くの？
	    Cell cell = row.createCell(0);

	    //入力したい値を書いて
	    cell.setCellValue("始めてのPOI");


	    //Excelファイルの保存先を決めるよ～
	    //生成用の変数を初期化して
	    FileOutputStream out = null;

	    //どこに保存したいかと末尾にファイルのお名前つけたげて～
	    String path = "C:\\Users\\onumaayano1199\\Pictures\\Sample2.xlsx";

	    try {
	      //この場所に生成するよ～
	      out = new FileOutputStream(path);

	      //ファイル作りまーす
	      wb.write(out);

	    } catch (Exception ex) {
	      ex.printStackTrace();

	    } finally {
	      try {
	        wb.close();
	        out.close();
	      } catch (Exception ex2) {
	        ex2.printStackTrace();
	      }
	    }
	  }

}
