package excle;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class ReadExcel
{
	public static void main(String[] args)
	{
		try
		{
			FileInputStream inputStream = new FileInputStream("read.xls");
			HSSFWorkbook hssfWorkbook = new HSSFWorkbook(inputStream);
			HSSFSheet sheet = hssfWorkbook.getSheetAt(0);
			Iterator<Row> rows = sheet.rowIterator();
			Map<String, MessageDTO> allMessages = new LinkedHashMap<>();
			while (rows.hasNext())
			{
				Row row = (Row)rows.next();
				//跳过第一行
				if(row.getRowNum() < 1)
					continue;
				Cell cell = row.getCell(0);
				Cell cellType = row.getCell(2);
				Cell cellPerson = row.getCell(1);
				String number = cell.getStringCellValue().substring(12, 17);
				if(allMessages.containsKey(number))
				{
					MessageDTO messageDTO = allMessages.get(number);
					if("原告".equals(cellType.getStringCellValue()))
					{
						String plainTiff = messageDTO.getPalinTiff();
						plainTiff += cellPerson.getStringCellValue() + " ";
						messageDTO.setPalinTiff(plainTiff);
					}
					else 
					{
						String defendant = messageDTO.getDefendant();
						defendant += cellPerson.getStringCellValue() + " ";
						messageDTO.setDefendant(defendant);
					}
				}
				else
				{
					MessageDTO messageDTO = new MessageDTO();
					messageDTO.setFileName(number);
					messageDTO.setNumber(number);
					if("原告".equals(cellType.getStringCellValue()))
						messageDTO.setPalinTiff(cellPerson.getStringCellValue() + " ");
					else {
						messageDTO.setDefendant(cellPerson.getStringCellValue() + " ");
					}
					allMessages.put(number, messageDTO);
				}
//				System.out.print(number + "  ");
			}
			hssfWorkbook.close();
			System.out.println("success read");
			
			for(MessageDTO messageDTO : allMessages.values())
			{
				exportBook(messageDTO.getFileName(), messageDTO.getNumber(), messageDTO.getPalinTiff(), messageDTO.getDefendant());
			}
			System.out.println("success export");
		} catch (Exception e)
		{
			// TODO: handle exception
			System.out.println(e.toString());
		}
		
	}
	
	public static void exportBook(String name,String number,String plainTiff,String defendant) throws Exception
	{
		Workbook workbook = Workbook.getWorkbook(new File("template.xls"));
		String fileName = "./export/" + name + ".xls";
		WritableWorkbook exportBook = Workbook.createWorkbook(new File(fileName), workbook);
		WritableSheet sheet = exportBook.getSheet("封面");
		Label label1 = (Label)sheet.getWritableCell(15,6);
		label1.setString(number);
		
//		Label label2 = (Label)sheet.getWritableCell(7,7);
//		label2.setString("有病把");
		
		Label label3 = (Label)sheet.getWritableCell(7,8);
		label3.setString(plainTiff);
		
		Label label = (Label)sheet.getWritableCell(7,10);
		label.setString(defendant);
		exportBook.write();
		exportBook.close();
		System.out.print("success " + name);
	}
}
