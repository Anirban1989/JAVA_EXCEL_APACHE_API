package com.soc.excel;

import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.CellStyle;

public class MainTest{
	public static void main(String[] args) {
		try {
			//ExcelModTest();
			ExcelReadTest();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// ExcelWriteTest();
	}

	/**
	 * 개발중 
	 * @throws IOException
	 */
	public static void ExcelModTest() throws IOException{
		String filePath = "d:\\TEST2.xlsx";
		ExcelUtil wb;
		wb = new ExcelUtil(filePath);
		CellStyle cellStyle = wb.getCellStyle();
		cellStyle.setBorderTop((short)1);
		cellStyle.setBorderLeft((short)1);
		cellStyle.setBorderRight((short)1);
		//wb.insertCellString(null, 2, 3, 3, "테스트 텍스트", cellStyle);
		//wb.insertImageFile(null, 3, 3, 3, 2, 3, null, "d:\\asd.jpg");
		
	}

	/**
	 *  엑셀 파일 읽기 테스트 예제 
	 *  엑셀로 부터 불러들인 정보를 ExcelArray 자료구조로 이루어진 ArrayList 에 담고 해당 정보를 엑셀 파일로 만들어내기 
	 * @throws IOException
	 */
	public static void ExcelReadTest() throws IOException{
		String filePath = "d:\\TEST2.xlsx";
		ArrayList<ExcelArray> resultData = new ArrayList<ExcelArray>();
		ExcelUtil wbload = new ExcelUtil(filePath);
		resultData = wbload.loadExcelArray(0);
		ExcelUtil wb = new ExcelUtil("d:\\2222saveTest.xlsx");
		for ( int i = 0;i < resultData.size();i++) {
			ExcelArray excelInfo = resultData.get(i);
			int cul = excelInfo.getExcelCulIndex();
			int row = excelInfo.getExcelRowIndex();
			int sheet = excelInfo.getExcelSheet();
			String cellText = excelInfo.getTextString();
			CellStyle cellStyle = wb.getCellStyle();
			System.out.println("sheet :" + sheet + "  row:" + row + "  cul:" + cul + "  text:" + cellText);
			cellStyle.setBorderLeft((short) 1);
			if (i > 5) {
				cellStyle.setBorderTop((short) 1);
			} else {
				cellStyle.setBorderTop((short) 0);
			}
			System.out.println("sheet :" + sheet + "  row:" + row + "  cul:" + cul + "  text:" + cellText);
			wb.insertCellString(null, sheet, row, cul, 512,7000,cellText, cellStyle);
		}
		try {
			wb.insertImageFile  (null, 2, 3, 3,880,7000, 2, 3, null, "d:\\asd.jpg");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		wb.saveExcelFile();
	}	
}
