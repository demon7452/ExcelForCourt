package excel;

import java.io.File;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class Read
{
	public static void main(String [] args)
	{
		try
		{
			
//			WritableWorkbook workbook2 = Workbook.createWorkbook(new File("tst2.xls"), workbook);
//			WritableSheet sheet = workbook2.getSheet("封面");
//			Label label1 = (Label)sheet.getWritableCell(15,6);
//			label1.setString("12345");
//			
////			Label label2 = (Label)sheet.getWritableCell(7,7);
////			label2.setString("有病把");
//			
//			Label label3 = (Label)sheet.getWritableCell(7,8);
//			label3.setString("别问我");
//			
//			Label label = (Label)sheet.getWritableCell(7,10);
//			label.setString("就是他2");
//			workbook2.write();
//			workbook2.close();
//			System.out.println("success");
//			Workbook workbook2 = Workbook.getWorkbook(new File("template.xls"));
			exportBook("a");
			exportBook("b");
		} catch (Exception e)
		{
			// TODO: handle exception
			System.out.println(e.toString());
		}
	}
	
	public static void exportBook(String name) throws Exception
	{
		Workbook workbook = Workbook.getWorkbook(new File("template.xls"));
		name += ".xls";
		WritableWorkbook exportBook = Workbook.createWorkbook(new File(name), workbook);
		WritableSheet sheet = exportBook.getSheet("封面");
		Label label1 = (Label)sheet.getWritableCell(15,6);
		label1.setString("12345");
		
//		Label label2 = (Label)sheet.getWritableCell(7,7);
//		label2.setString("有病把");
		
		Label label3 = (Label)sheet.getWritableCell(7,8);
		label3.setString("别问我");
		
		Label label = (Label)sheet.getWritableCell(7,10);
		label.setString("就是他2");
		exportBook.write();
		exportBook.close();
		System.out.println("success " + name);
	}
}
