/*********************************************************************
 * CLASS NAME: ExcelUtil                                        
 * © Copyright IBM Corporation 2016. All rights reserved.            
 * 
 * CHANGE ACTIVITY:
 * DATE     | AUTHOR     | ACTIVITY                    
 *----------+------------+-------------------------------------------
 * 04/22/16 | TCOE       | Initial Version          
 *----------+------------+-------------------------------------------
 *          |            |                                           
 ********************************************************************/
package util;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class ExcelUtil {

	private static Workbook workbook = null;
	private static WritableWorkbook newbook = null;
	private static Sheet sheet = null;
	private static String excelName = null;
	private static File excelFile = null;
	private static int sheetNum = 0;
	private static int rowNum = 0;
	private static int columNum = 0;
	private static ArrayList<String> headerList = new ArrayList<String>();
	//private static ArrayList<HashMap<String, HashMap<String, String[]>>> testCaseList = new ArrayList<HashMap<String, HashMap<String, String[]>>>();

	/*---------------------------------------------------------------------------
	 * Class Constructor
	 *-------------------------------------------------------------------------*/
	public ExcelUtil(String filename, int sheetnum) {
		excelName = filename;
		excelFile = new File(excelName);
		sheetNum = sheetnum;
	}

	public ExcelUtil(File filepath, int sheetnum) {
		excelFile = filepath;
		sheetNum = sheetnum;
	}

	public ExcelUtil(String filename) {
		excelName = filename;
		excelFile = new File(excelName);
		sheetNum = 0;
	}

	public ExcelUtil(File filepath) {
		excelFile = filepath;
		sheetNum = 0;
	}

	/*--------------------------------------------------------------------------- 
	 * Get class attributes 
	 *---------------------------------------------------------------------------*/
	public ArrayList<String> ReturnHeader() {
		return headerList;
	}

	public int GetRowNum() {
		OpenExcel();
		workbook.close();
		return rowNum;
	}

	public int GetColumnNum() {
		OpenExcel();
		workbook.close();
		return columNum;
	}

	/*---------------------------------------------------------------------------
	 * Open an exist excel sheet
	 *-------------------------------------------------------------------------*/
	private void OpenExcel() {
		try {
			workbook = Workbook.getWorkbook(excelFile);
			sheet = workbook.getSheet(sheetNum);
			columNum = sheet.getColumns(); // get column number
			rowNum = sheet.getRows(); // get row number
			// System.out.println("Row number: " + rowNum);
			// System.out.println("Column number: " + columNum);
		} catch (BiffException | IOException e) {
			System.out.println(e);
			e.printStackTrace();
		}
	}

	/*---------------------------------------------------------------------------
	 * Create new excel sheet
	 *-------------------------------------------------------------------------*/
	public static void CreateExcel(String filename, String sheetname,
			int sheetnum, HashMap<String, String> datamap) {
		try {
			// Create new excel file
			newbook = Workbook.createWorkbook(new File(filename));
			// Create first sheet with the parameter 0
			WritableSheet sheet = newbook.createSheet(sheetname, sheetnum);

			Label label;
			int i = 0;
			for (String keys : datamap.keySet()) {
				// Put the data into the cell(column,row)
				label = new Label(i, 0, keys);
				// Add label to sheet
				sheet.addCell(label);
				label = new Label(i, 1, datamap.get(keys));
				sheet.addCell(label);
				sheet.setColumnView(i, 20);
				i++;
			}

			sheet.setRowView(0, 300);

			// Write data to excel and close file
			newbook.write();
			newbook.close();

		} catch (Exception e) {
			System.out.println(e);
		}
	}

	/*---------------------------------------------------------------------------
	 * Read excel sheet
	 *-------------------------------------------------------------------------*/
	// Read excel and print all the data in it
	public void ReadExcelShowAllData() {
		OpenExcel();
		for (int i = 0; i < rowNum; i++) // 循环进行读写
		{
			for (int j = 0; j < columNum; j++) {
				Cell cell1 = sheet.getCell(j, i);
				String result = cell1.getContents();
				System.out.print(result);
				System.out.print(" \t ");
			}
			System.out.println();
		}
		workbook.close();
	}

	// Read excel case by the case ID
	public HashMap<String, String> ReadCaseByCaseID(String caseid) {
		HashMap<String, String> caseMap = new HashMap<String, String>();
		headerList.clear();

		OpenExcel();

		// Search for case by case ID
		for (int i = 0; i < rowNum; i++) {
			Cell cellsearch = sheet.getCell(0, i);
			String result1 = cellsearch.getContents();
			if (result1.equals(caseid)) {
				System.out.println("Case ID Search Result: " + result1 + "\n");
				for (int j = 0; j < columNum; j++) {
					Cell cellHeader = sheet.getCell(j, 0);
					Cell cellItem = sheet.getCell(j, i);
					String header = cellHeader.getContents();
					String item = cellItem.getContents();

					headerList.add(j, header);
					caseMap.put(header, (item.isEmpty()) ? "N/A" : item);
				}
				break;
			}
		}
		// Print search result
		// for (String keys : headerList) {
		// System.out.println(keys + " : \n\t\t" + caseMap.get(keys));
		// }

		workbook.close();
		return caseMap;
	}

	// Read excel case by the line number
	public HashMap<String, String> ReadCaseByRow(int rowNum) {
		HashMap<String, String> caseMap = new HashMap<String, String>();
		caseMap.clear();

		OpenExcel();

		for (int j = 0; j < columNum; j++) {
			Cell cellHeader = sheet.getCell(j, 0);
			Cell cellItem = sheet.getCell(j, rowNum);
			String header = cellHeader.getContents();
			String item = cellItem.getContents();

			caseMap.put(header, (item.isEmpty()) ? "N/A" : item);

		}
		// for (String keys : caseMap.keySet()) {
		// System.out.println(keys + " : \n\t\t" + caseMap.get(keys));
		// }

		workbook.close();
		workbook.close();
		return caseMap;
	}

	// Read excel by its header name and list all item below this header
	public void ReadCaseByColumn(String headername) {
		OpenExcel();

		for (int i = 0; i < columNum; i++) {
			Cell cell1 = sheet.getCell(i, 0);
			String result1 = cell1.getContents();
			if (result1.equals(headername)) {
				System.out.println("Header Name: " + result1);
				for (int j = 1; j < rowNum; j++) {
					Cell cell2 = sheet.getCell(i, j);
					String result2 = cell2.getContents();
					System.out.print(result2);
					System.out.print(" \t ");
				}
			}
			System.out.println();
		}
		workbook.close();
	}

	public static ArrayList<HashMap<String, String[]>> getTestCase(
			String excelpath, String sheetname, String subjectName) {

		// List<String> header = null;
		ArrayList<HashMap<String, String[]>> testCaseList = new ArrayList<HashMap<String, String[]>>();
		HashMap<String, String[]> content;
		
		//Map<String, ArrayList<String[]>> rowmap = new HashMap<String, ArrayList<String[]>>();
		boolean findrow = false;
		int beginrownumber = 0;
		int endrownumber = 0;		

		try {
			Workbook workbook = Workbook.getWorkbook(new File(excelpath));
			Sheet sheet = workbook.getSheet(sheetname);
			// the row index begin with 0
			int rows = sheet.getRows();
			// the column index begin with 0
			int columns = sheet.getColumns();
			
			//System.out.println("Sheet: " + sheetname + " Total rows number: " + rows);//TODO  Delete Test statement
			for (int rowindex = 1; rowindex < rows; rowindex++) {
				String cellcontent = sheet.getCell(0, rowindex).getContents()
						.toLowerCase().trim();
				if (cellcontent.equalsIgnoreCase(subjectName)) {
					if (!findrow) {
						beginrownumber = rowindex;
						findrow = true;
					}
				} else if (cellcontent.equalsIgnoreCase("") && rowindex != rows-1) {
					continue;
				} else if (!cellcontent.equalsIgnoreCase("") && !cellcontent.equalsIgnoreCase(subjectName)) {
					endrownumber = rowindex - 1;
					break;
				} else if (cellcontent.equalsIgnoreCase("") && rowindex == rows-1){
						endrownumber = rowindex;
						break;
				}
			}
			//System.out.println(subjectName + " begin row number: " + beginrownumber);//TODO  Delete Test statement
			//System.out.println(subjectName + " end row number: " + endrownumber);//TODO  Delete Test statement

			int eachCaseBeginRow = beginrownumber;
			int eachCaseEndRow = beginrownumber;
						
			if (findrow) {
				String[] temp;
				for (int rowindex = 1; rowindex <= endrownumber; rowindex++) {
					String casenameCheck = sheet.getCell(1, rowindex).getContents();
					if(!casenameCheck.equalsIgnoreCase("")){
						eachCaseBeginRow = rowindex;
						eachCaseEndRow = rowindex;						
						//System.out.println("Case name " + casenameCheck);//TODO  Delete Test statement
					}else if(casenameCheck.equalsIgnoreCase("") 
							&& rowindex != endrownumber 
							&& sheet.getCell(1, rowindex+1).getContents().equals("")){
						continue;				
					}else if(rowindex == endrownumber || !sheet.getCell(1, rowindex+1).getContents().equals("")){
						eachCaseEndRow = rowindex;					
						
						//System.out.println("Current case begin at: " + eachCaseBeginRow); //TODO Delete Test statement
						//System.out.println("Current case end at: " + eachCaseEndRow); //TODO  Delete Test statement
						
						content = new HashMap<String, String[]>();
						int caseStepNum = eachCaseEndRow - eachCaseBeginRow + 1;
						// Extract value
						for (int columnindex = 5; columnindex < columns; columnindex++) {
							String currentHeader = sheet.getCell(columnindex, 0).getContents();
							
							if(currentHeader.equals("Description")){
								temp = new String[1];
								temp[0] = sheet.getCell(columnindex, eachCaseBeginRow).getContents();
								content.put("Description", temp);								
							}else if(currentHeader.equals("Step Description")){
								temp = new String[caseStepNum];
								for(int index = 0;index < caseStepNum;index++){
									temp[index] = sheet.getCell(columnindex, eachCaseBeginRow + index).getContents();
								}								
								content.put("Step Description", temp);
							}else if(currentHeader.equals("Expected Results")){
								temp = new String[caseStepNum];
								for(int index = 0;index < caseStepNum;index++){
									temp[index] = sheet.getCell(columnindex, eachCaseBeginRow + index).getContents();
								}
								content.put("Expected Results", temp);
							}else if(currentHeader.equals("Actual Results")){
								temp = new String[caseStepNum];
								for(int index = 0;index < caseStepNum;index++){
									temp[index] = sheet.getCell(columnindex, eachCaseBeginRow + index).getContents();
								}
								content.put("Actual Results", temp);
							}else if(currentHeader.equals("Screen Shot")){
								temp = new String[caseStepNum];
								for(int index = 0;index < caseStepNum;index++){
									temp[index] = sheet.getCell(columnindex, eachCaseBeginRow + index).getContents();
								}
								content.put("Screen Shot", temp);
							}else if(currentHeader.equals("Website")){
								temp = new String[1];
								temp[0] = sheet.getCell(columnindex, eachCaseBeginRow).getContents();
								content.put("Website", temp);								
							}else if(currentHeader.equals("Browser")){
								temp = new String[1];
								temp[0] = sheet.getCell(columnindex, eachCaseBeginRow).getContents();
								content.put("Browser", temp);								
							}else if(currentHeader.equals("Prerequisites")){
								temp = new String[1];
								temp[0] = sheet.getCell(columnindex, eachCaseBeginRow).getContents();
								content.put("Prerequisites", temp);								
							}														
						}
						
						String CurrentCasename = sheet.getCell(1, eachCaseBeginRow).getContents();
						temp = new String[1];
						temp[0] = CurrentCasename;
						content.put("Test Name", temp);	
						
						testCaseList.add(content);
					}				
				}
			}

		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return testCaseList;
	}

}
