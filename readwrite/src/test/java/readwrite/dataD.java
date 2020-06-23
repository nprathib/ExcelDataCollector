package readwrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataD {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		FileInputStream fis = new FileInputStream("E://demoData//A09B4D10.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		int sheets = workbook.getNumberOfSheets();
		System.out.println("no. of sheets"+ sheets);
		String columnData;
		for(int i=0;i<sheets;i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("work"))
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
					if(value.getStringCellValue().equalsIgnoreCase("role"))
					{
						column=k;
						columnData = getCellData(firstrow,column);
						if(columnData==null)
						{
							break;
						}
						System.out.println("cell value: "+value.getStringCellValue());
						break;
					}
					k++;
				}

				while(rows.hasNext())
				{
					Row r = rows.next();
					int rn = r.getRowNum();
					//System.out.println("value of rn"+ rn);
					int rv=0;
					int j=0;
					for(int m=0;m<=rn;m++)
					{
						//System.out.println(r.getCell(column).getStringCellValue());
						
						
					 columnData = getCellData(r,column);
						if(columnData==null)
						{
							break;
						}
					if(columnData.equalsIgnoreCase("national"))
					//if(r.getCell(column).getStringCellValue().equalsIgnoreCase("rcm"))
					{
						//System.out.println("is it in");
						java.util.Iterator<Cell> cv = r.cellIterator();
						int columnn1 =0;
						while(cv.hasNext())
						{
							if(cv.hasNext()) {
							columnData = getCellData(r,columnn1);
							if(columnData==null)
							{
								break;
							}
							//System.out.println(r.getCell(columnn1).getStringCellValue());
							
							}
							else {break;}
							columnn1++;
						}
						rv=j;
						
					}
					/*
					 * else { System.out.println(""); break; }
					 */
					j++;
					//System.out.println("value of rv"+ rv);
					
				}
		
			}
		}
	}
}
	
	public static String getCellData(Row r, int column)  {
		String columnData = null;
		try {
		try {
			 columnData = r.getCell(column).getStringCellValue();
			 System.out.println(r.getCell(column).getStringCellValue());
		}
		catch(Exception e){
			double numericcolumnData = r.getCell(column).getNumericCellValue();
			columnData = Double. toString(numericcolumnData);
			System.out.println((int)Math.round(r.getCell(column).getNumericCellValue()));
		}
		}
	 catch(NullPointerException e) {
	 System.out.println("Caught null pointer exception");
	 
	 }
		
	
	//if(columnData.equalsIgnoreCase("rcm"))
		return columnData;
	}
	
}
			
		

	


