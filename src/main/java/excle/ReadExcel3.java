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

public class ReadExcel3
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
				Cell cell = row.getCell(1);
				Cell cellPerson = row.getCell(2);
				Cell cellReson = row.getCell(3);
				Cell cellCloseWay = row.getCell(10);
				Cell cellJudge = row.getCell(5);
				Cell cellChiefJudge = row.getCell(6);
				Cell cellRecordDate = row.getCell(4);
				Cell cellEndDate = row.getCell(9);
				Cell cellCollegiateMembers = row.getCell(8);
				String number = cell.getStringCellValue().substring(12, 17);
				MessageDTO messageDTO = new MessageDTO();
				messageDTO.setFileName(number);
				messageDTO.setNumber(number);
				messageDTO.setCaseReson(cellReson.getStringCellValue());
				int index = cellPerson.getStringCellValue().indexOf(";");
				messageDTO.setPalinTiff(cellPerson.getStringCellValue().substring(3,index).replace(',', ' '));
				messageDTO.setDefendant(cellPerson.getStringCellValue().substring(index+4,cellPerson.getStringCellValue().length()).replace(',', ' '));
				messageDTO.setCloseWay(cellCloseWay.getStringCellValue());
				messageDTO.setJudge(cellJudge==null?"":cellJudge.getStringCellValue());
				messageDTO.setChiefJudge(cellChiefJudge.getStringCellValue());
				messageDTO.setRecordDate(cellRecordDate.getStringCellValue());
				messageDTO.setEndDate(cellEndDate.getStringCellValue());
				messageDTO.setCollegiateMembers(cellCollegiateMembers.getStringCellValue());
				allMessages.put(number, messageDTO);
			}
			hssfWorkbook.close();
			System.out.println("success read");
			int num = 0;
			for(MessageDTO messageDTO : allMessages.values())
			{
				exportBook(messageDTO);
				num ++;
				if(num % 10 == 0)
					System.out.println("");
			}
			System.out.println("");
			System.out.println("success export " + num);
		} catch (Exception e)
		{
			// TODO: handle exception
			System.out.println(e.toString());
		}
		
	}
	
	public static void exportBook(MessageDTO messageDTO) throws Exception
	{
		String name = messageDTO.getFileName();
		String number = messageDTO.getNumber();
		String caseReson = messageDTO.getCaseReson();
		String plainTiff = messageDTO.getPalinTiff();
		String defendant = messageDTO.getDefendant();
		String closeWay = messageDTO.getCloseWay();
		String judge = messageDTO.getJudge();
		String chiefJudge = messageDTO.getChiefJudge();
		String[] collegiateMembers = messageDTO.getCollegiateMembers().split("、");
		
		Workbook workbook = Workbook.getWorkbook(new File("template.xls"));
		String fileName = "./export/" + name + ".xls";
		WritableWorkbook exportBook = Workbook.createWorkbook(new File(fileName), workbook);
		WritableSheet sheet = exportBook.getSheet("封面");
		Label label1 = (Label)sheet.getWritableCell(15,6);
		label1.setString(number);
		
		Label label2 = (Label)sheet.getWritableCell(7,7);
		label2.setString(caseReson);
		
		Label label3 = (Label)sheet.getWritableCell(7,8);
		label3.setString(plainTiff);
		
		Label label = (Label)sheet.getWritableCell(7,10);
		label.setString(defendant);
		
		Label label5 = (Label)sheet.getWritableCell(4,15);
		label5.setString(closeWay);
		
		Label labelJudgeType = (Label)sheet.getWritableCell(3,11);
		Label labelJudgePerson = (Label)sheet.getWritableCell(3,12);
		Label labelCollegiateMember1 = (Label)sheet.getWritableCell(7,12);
		Label labelCollegiateMember2 = (Label)sheet.getWritableCell(11,12);
		if("".equals(chiefJudge))
		{
			labelJudgeType.setString("审判员");
			labelJudgePerson.setString(judge);
			labelCollegiateMember1.setString("");
			labelCollegiateMember2.setString("");
		}
		else
		{
			labelJudgeType.setString("审判长");
			labelJudgePerson.setString(chiefJudge);
			labelCollegiateMember1.setString(collegiateMembers[0]);
			labelCollegiateMember2.setString(collegiateMembers[1]);
		}
		Label labelRecordYear = (Label)sheet.getWritableCell(4,13);
		labelRecordYear.setString(messageDTO.getRecordDate().substring(0,4));
		Label labelRecordMonth = (Label)sheet.getWritableCell(7,13);
		labelRecordMonth.setString(messageDTO.getRecordDate().substring(5,7));
		Label labelRecordDay = (Label)sheet.getWritableCell(9,13);
		labelRecordDay.setString(messageDTO.getRecordDate().substring(8,10));
		
		String endDate = messageDTO.getEndDate();
		Label labelEndYear = (Label)sheet.getWritableCell(12,13);
		Label labelEndMonth = (Label)sheet.getWritableCell(15,13);
		Label labelEndDay = (Label)sheet.getWritableCell(18,13);
		if("".equals(endDate))
		{
			labelEndYear.setString("");
			labelEndMonth.setString("");
			labelEndDay.setString("");
		}
		else
		{
			labelEndYear.setString(endDate.substring(0, 4));
			labelEndMonth.setString(endDate.substring(5,7));
			labelEndDay.setString(endDate.substring(8,10));
		}
		
		
		exportBook.write();
		exportBook.close();
		System.out.print("success " + name + "  ");
	}
}
