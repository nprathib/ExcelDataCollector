package readwrite;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataD1 {
	

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		FileInputStream fis = new FileInputStream("E://demoData//A09B4D10.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		int sheets = workbook.getNumberOfSheets();
		System.out.println("no. of sheets"+ sheets);
		
		for(int i=0;i<sheets;i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("acces"))
			{
				XSSFSheet sheet = workbook.getSheetAt(i);
				System.out.println("could fetch the sheet");
				java.util.Iterator<Row> rows = sheet.iterator();
				Row firstrow=rows.next();
				java.util.Iterator<Cell> ce = firstrow.cellIterator();
				int k=0;
				int column = 0;
				while(ce.hasNext())
				{
					Cell value = ce.next();
					if(value.getStringCellValue().equalsIgnoreCase("app"))
					{
						column=k;
						System.out.println("able to fetch cell value");
						break;
					}
					k++;
				}
			System.out.println(column);
				while(rows.hasNext())
				{
					
					Row r=rows.next();
					System.out.println("is it in");
					
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase("qwe"))
					{
						
						////after you grab purchase testcase row = pull all the data of that row and feed into test
						
						java.util.Iterator<Cell>  cv=r.cellIterator();
						while(cv.hasNext())
						{
						Cell c=	cv.next();
					System.out.println(c.getStringCellValue());
						
						}
					}
					
				
				}
		
			}
		}
	}
}

