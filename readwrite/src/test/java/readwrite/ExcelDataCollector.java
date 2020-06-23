package readwrite;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelDataCollector {

	static int k=0;
	static int columnNumber = 0;
	static String columnData = null;
	static String rowandColumnData = null;

		public static void main(String[] args) throws IOException {
			// TODO Auto-generated method stub
			FileInputStream fis = new FileInputStream("E://demoData//A09B4D10.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			
			int sheets = workbook.getNumberOfSheets();
			System.out.println("no. of sheets"+ sheets);
			

//below loop is to fetch the data from Worksheet			
			for(int i=0;i<sheets;i++)
			{
				if(workbook.getSheetName(i).equalsIgnoreCase("work"))
				{
					XSSFSheet sheet = workbook.getSheetAt(i);
					System.out.println("could fetch the sheet");
					
					java.util.Iterator<Row> rows = sheet.iterator();
					Row firstrow=rows.next();
					java.util.Iterator<Cell> cell = firstrow.cellIterator();
					
				    readFirstRowForColumnNumber(firstrow,cell);//read the first row data
					//columnNumber
					//System.out.println(firstrow+""+cell);
					
					readColumnData(rows, columnNumber);//read the column data for specific row and column
					//System.out.println(rows+""+column);
					
					
			}
		}
	}
		//below method is used to get the complete row value for provided header and column data
		public static void readColumnData(java.util.Iterator<Row> rows,int column) {
			
		 while(rows.hasNext()) {
				Row r = rows.next();
				//int rn = r.getRowNum();
					
				columnData = getCellData(r,column);
					//System.out.println("Verticle Column Data" + columnData);
					
					if(columnData.equalsIgnoreCase("national"))
					{
						System.out.println(columnData+ " -- Row Data From Excel");
						java.util.Iterator<Cell> cv = r.cellIterator();
						int columnn1 =0;
						
						while(cv.hasNext())				//This while is to get data of complete row
						{
							if(cv.hasNext()) {
							rowandColumnData = getCellData(r,columnn1);
							//System.out.println(r+""+columnn1);
							System.out.println("Entire Row value From Excel: "+rowandColumnData);
							if(rowandColumnData==null)
							{
								break;
							}
						}
							columnn1++;
						}
					}
		 		}
			}
		//below method is used to find the cell value from first row

		public static int readFirstRowForColumnNumber(Row firstrow,java.util.Iterator<Cell> cell) {
			
			while(cell.hasNext())
			{
				Cell value = cell.next();
				if(value.getStringCellValue().equalsIgnoreCase("role"))
				{
					columnNumber=k;
					System.out.println("Cell value: "+value.getStringCellValue());
					break;
				}
				k++;
			}
			return columnNumber;
		}
		//below method is used to get the cell data from excel

		public static String getCellData(Row r, int column)  {
			String columnDataVal = null;
			try {
				try {
					 columnDataVal = r.getCell(column).getStringCellValue();
					// System.out.println(r.getCell(column).getStringCellValue());
				}
				catch(Exception e){
					double numericcolumnData =r.getCell(column).getNumericCellValue();
					int numbercolumnData = (int) Math.round(numericcolumnData);
					columnDataVal = Double.toString(numbercolumnData).replace(".0", "");
					//System.out.println((int)Math.round(r.getCell(column).getNumericCellValue()));
				}
			}catch(NullPointerException e) { }
				return columnDataVal;
			}
		
	}
				
			

		


