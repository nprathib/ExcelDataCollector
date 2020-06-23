package readwrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

import javax.swing.text.html.HTMLDocument.Iterator;

import org.apache.poi.ss.format.CellTextFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {
	
	//Identify Testcases coloum by scanning the entire 1st row
		//once coloumn is identified then scan entire testcase coloum to identify purcjhase testcase row
		//after you grab purchase testcase row = pull all the data of that row and feed into test
		
		public ArrayList getData(String testcaseName) throws IOException
		{
			//fileInputStream argument
					ArrayList<String> a=new ArrayList<String>();
					
					FileInputStream fis=new FileInputStream("E://demoData//A09B4D10.xlsx");
					XSSFWorkbook workbook=new XSSFWorkbook(fis);
					
					int sheets=workbook.getNumberOfSheets();
					for(int i=0;i<sheets;i++)
					{
						if(workbook.getSheetName(i).equalsIgnoreCase("acces"))
								{
						XSSFSheet sheet=workbook.getSheetAt(i);
						//Identify Testcases coloum by scanning the entire 1st row
						System.out.println("know the sheet"+" "+sheet);
						
						 java.util.Iterator<Row>  rows= sheet.iterator();// sheet is collection of rows
						Row firstrow= rows.next();
						java.util.Iterator<Cell> ce=firstrow.cellIterator();//row is collection of cells
						int k=0;
						int coloumn = 0;
					while(ce.hasNext())
					{
						Cell value=ce.next();
						
						if(value.getStringCellValue().equalsIgnoreCase("user"))
						{
							coloumn=k;
							System.out.println("know the column");
							
						}
						
						k++;
					}
					System.out.println(coloumn);
					
					////once coloumn is identified then scan entire testcase coloum to identify purcjhase testcase row
					while(rows.hasNext())
					{
						
						Row r=rows.next();
						System.out.println("know the row");
						
						if(r.getCell(coloumn).getStringCellValue().equalsIgnoreCase(testcaseName))
						{
							
							////after you grab purchase testcase row = pull all the data of that row and feed into test
							System.out.println("i am in");
							java.util.Iterator<Cell>  cv=r.cellIterator();
							while(cv.hasNext())
							{
							Cell c=	cv.next();
							
								
							a.add(c.getStringCellValue());
							
							}
						}
						
					
					}
				
					
						
						
						
						
						
						
						
								}
					}
					return a;
					
		}

		public static void main(String[] args) throws IOException {
			// TODO Auto-generated method stub

			
		}


	}
		
	

