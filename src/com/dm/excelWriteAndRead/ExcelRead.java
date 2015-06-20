package com.dm.excelWriteAndRead;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import org.junit.Test;

public class ExcelRead {
	@Test
	public void read(){
		Workbook readwb = null;
		try {
			InputStream instream = new FileInputStream("e:/excelFile/红楼梦.xls");
			readwb = Workbook.getWorkbook(instream);
			
			//获取第一张sheeet表
			Sheet readSheet = readwb.getSheet(0);
			//获取总列数
			int rsColumns = readSheet.getColumns();
			//获取总行数
			int rsRows = readSheet.getRows();
			
			for(int i=0;i<rsRows;i++){
				for(int j=0;j<rsColumns;j++){
					Cell cell = readSheet.getCell(j, i);
					System.out.print(cell.getContents()+" :");
				}
				System.out.println();
			}
	
			WritableWorkbook wwb = Workbook.createWorkbook(new File("e:/excelFile/红楼梦1.xls"), readwb);
			WritableSheet ws= wwb.getSheet(0);
			WritableCell wc = ws.getWritableCell(1, 0);
			if(wc.getType() == CellType.LABEL){
				Label l = (Label)wc;
				l.setString("新姓名");
			}
			wwb.write();
			wwb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			readwb.close();
		}
	}
	
	@Test
	public void testWrite(){
		WritableWorkbook book = null;
		try {
			book = Workbook.createWorkbook(new File("e:/excelFile/测试.xls"));
			WritableSheet sheet = book.createSheet("第一页",0);
			Label label = new Label(0, 0, "测试");
			
			WritableFont wfc = new WritableFont(WritableFont.ARIAL,10,WritableFont.NO_BOLD,false,UnderlineStyle.NO_UNDERLINE,jxl.format.Colour.DARK_YELLOW); 
			WritableCellFormat wcfFC = new WritableCellFormat(wfc);
			sheet.addCell(label);
			
			Number number = new Number(1, 0, 123.456,wcfFC);
			sheet.addCell(number);
			
			Label s = new Label(1, 2, "三十三");
			sheet.addCell(s);
			book.write();
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			try {
				book.close();
			} catch (WriteException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	@Test
	public void testWritePaperless(){
		String[] options = {"A","B","C","D","E"};
		WritableWorkbook book = null;
		int count = 20;
		
		try {
			WritableFont fontTitle = new WritableFont(WritableFont.ARIAL,9,WritableFont.BOLD,false,UnderlineStyle.NO_UNDERLINE,jxl.format.Colour.BLACK);  
			fontTitle.setColour(jxl.format.Colour.RED);  
			WritableCellFormat formatTitle = new WritableCellFormat(fontTitle);  
			formatTitle.setAlignment(Alignment.CENTRE);
			
//			CellView navCellView = new CellView();  
//		    navCellView.setAutosize(true); //设置自动大小
//		    navCellView.setSize(28);
			
			book = Workbook.createWorkbook(new File("e:/excelFile/测试.xls"));
			WritableSheet sheet = book.createSheet("单选题", 0);
			Label number = new Label(0, 0, "序号",formatTitle);
			Label context = new Label(1, 0, "题干",formatTitle);
			Label option = new Label(2, 0, "选项",formatTitle);
			
			sheet.setColumnView(1, 20);
			
			sheet.addCell(number);
			sheet.addCell(context);
			sheet.addCell(option);
			
			for(int i=1;i<count+1;i++){//行
				Number serial = new Number(0, 1+(i-1)*options.length, i);
				Label cont = new Label(1, 1+(i-1)*options.length,"我是单选题的第"+i+"题");
				sheet.addCell(serial);
				sheet.addCell(cont);
			}
			
			for(int k=0;k<(count)*options.length;k++){
					Label opt = new Label(2, k+1, options[k%options.length]);
					sheet.addCell(opt);
			}
			
			book.write();
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			try {
				book.close();
			} catch (WriteException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
