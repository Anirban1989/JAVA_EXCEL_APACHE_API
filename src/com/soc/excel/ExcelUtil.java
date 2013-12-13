package com.soc.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** 엑셀 파일 입출력 Class
 * LastUpdate 2013-05-10 
 * -loadExcelArray(int):ExcelArray에 담은 엑셀내용 불러오기 함수
 * -getCellStyle():내부에 생성된 WorkBook의 상속을 받는 CellStyle불러오기
 * -insertCellString(String,int, int, int, String, CellStyle):엑셀에 스트링 값을 입력하는 함수
 * -insertImageFile(String, int, int, int, int, int, CellStyle,String):엑셀에 그림을
 * 넣는 함수 -saveExcelFile():입력이 끝난 엑셀을 최종적으로 저장하는 함수
 * 
 * @author zuneho 
 */
public class ExcelUtil{
	Workbook wb;
	Sheet sheet;
	Row row;
	Cell cell;
	CellStyle cellStyle;
	String filePath;
	String fileType;

	/**
	 * 생성자 초기 생성시 사용하며, XLS,XLSX 모두 지원한다.
	 * 
	 * @param filePath
	 *            엑셀파일이 존재하는 파일 경로를 함께 입력해 준다.(XLS, XLSX 모두 확장자를 포함하여 초기 생성한다)
	 */
	public ExcelUtil(String filePath) {
		this.filePath = filePath;
		String fileExtension = filePath.substring(filePath.lastIndexOf(".") + 1, filePath.length());
		if (fileExtension.toLowerCase().equals("xls")) {
			wb = new HSSFWorkbook();
			fileType = "xls";
		} else if (fileExtension.toLowerCase().equals("xlsx")) {
			wb = new XSSFWorkbook();
			fileType = "xlsx";
		} else {
			wb = null;
		}
	}

	/**
	 * 엑셀 파일의 빈곳을 제외한 모든 값들를 ExcelArray에 담아서 Sheet, row, cul의 값을 기억하여 저장한
	 * ArrayList를 리턴한다.
	 * 
	 * @param filePath
	 *            파일이 존재하는 실제 경로를 입력해준다.
	 * @param targetSheet
	 *            디폴트값은 0으로 설정하면 모든 시트를 읽는다.특정 시트만 읽을 경우 해당 시트값을 준다. (첫번째 시트만 읽고
	 *            싶다면 1, 3번째 시트를 읽고싶다면 3을 입력한다.)
	 * @return ExcelArray Class 의 각각의 cell 값을 담고 있는 ArrayList로 리턴한다.
	 */
	public ArrayList<ExcelArray> loadExcelArray(int targetSheet) {
		ArrayList<ExcelArray> excelArray = new ArrayList<ExcelArray>();
		if (wb != null) {
			try {
				// File file = new File(filePath);
				// fs = new POIFSFileSystem(new FileInputStream(file));
				if (fileType.equals("xls")) {
					wb = new HSSFWorkbook(new FileInputStream(filePath));
				} else if (fileType.equals("xlsx")) {
					wb = new XSSFWorkbook(new FileInputStream(filePath));
				}
				int sheetCount = wb.getNumberOfSheets();
				for ( int s = 0;s < sheetCount;s++) {
					// 시트 가져오기
					sheet = wb.getSheetAt(s);
					if (targetSheet > 0) {
						sheet = wb.getSheetAt(targetSheet - 1);
						s = targetSheet - 1;
					}
					// Row 갯수 가져오기
					int rows = sheet.getLastRowNum() + 1;
					for ( int r = 0;r < rows;r++) {
						if (sheet.getRow(r) != null) {
							// row 가져오기
							row = sheet.getRow(r);
							int cells = sheet.getRow(r).getLastCellNum();
							for ( int c = 0;c < cells;c++) {
								// cell 가져오기
								cell = row.getCell(c);
								String value = "";
								if (cell != null) {
									// cell 타입에 따른 데이타 처리
									switch (cell.getCellType()) {
									case Cell.CELL_TYPE_FORMULA:
										value = cell.getCellFormula();
										break;
									case Cell.CELL_TYPE_NUMERIC:
										// 더블형으로 사용할 때만 해당 주석을 변경한다.
										// value = ""
										// +cell.getNumericCellValue();
										long longvalue = (long) cell.getNumericCellValue();
										value = "" + String.valueOf(longvalue);
										break;
									case Cell.CELL_TYPE_STRING:
										value = "" + cell.getStringCellValue();
										break;
									case Cell.CELL_TYPE_BLANK:
										// value=""+cell.getBooleanCellValue();
										value = "";
										break;
									case Cell.CELL_TYPE_ERROR:
										value = "" + cell.getErrorCellValue();
										// value = "";
										break;
									default:
									}
								}
								if (value.length() > 0) {
									ExcelArray excel = new ExcelArray();
									excel.setExcelCulIndex(c);
									excel.setExcelRowIndex(r);
									excel.setExcelSheet(s);
									excel.setTextString(value);
									excelArray.add(excel);
								}
							}
						}
					}
					if (targetSheet > 0) {
						break;
					}
				}
			} catch (FileNotFoundException e) {
				ExcelArray excel = new ExcelArray();
				excel.setExcelCulIndex(0);
				excel.setExcelRowIndex(0);
				excel.setExcelSheet(0);
				excel.setTextString("File Not Found");
				excelArray.add(excel);
				return excelArray;
			} catch (IOException e) {
				ExcelArray excel = new ExcelArray();
				excel.setExcelCulIndex(0);
				excel.setExcelRowIndex(0);
				excel.setExcelSheet(0);
				excel.setTextString("IO Exception");
				excelArray.add(excel);
				return excelArray;
			}
		} else {
			ExcelArray excel = new ExcelArray();
			excel.setExcelCulIndex(0);
			excel.setExcelRowIndex(0);
			excel.setExcelSheet(0);
			excel.setTextString("Your Excel File destination can`t finde File !");
			excelArray.add(excel);
		}
		return excelArray;
	}

	/**
	 * ExcelUtil Class 내부에 생성된 WorkBook에 존재하는 CellStyle의 객체를 가져온다.
	 * 
	 * @return CellStyle (POI 내부 Class)
	 */
	public CellStyle getCellStyle() {
		if (wb != null) {
			cellStyle = wb.createCellStyle();
		}
		return cellStyle;
	}

	/**
	 * 엑셀 파일 내부에 텍스트를 입력한다.
	 * 
	 * @param sheetName
	 *            null 입력하려고 하는 목적 시트의 이름(새로 생성되는 시트만 이름을 바꿀 수 있다.)
	 * @param sheetIdx
	 *            not null 입력하려고 하는 목적 시트의 인데스
	 * @param rowIdx
	 *            not null 입력하려고 하는 목적 줄
	 * @param cellIdx
	 *            not null 입력하려고 하는 목적 열
	 * @param cellString
	 *            not null 입력하려고 하는 스트링 텍스트
	 * @param cellStyle
	 *            null 입력하려고 하는 Cell의 고유 스타일
	 * @return boolean 작업 성공여부
	 */
	public boolean insertCellString(String sheetName, int sheetIdx, int rowIdx, int cellIdx, int rowHeight,int cellSize,String cellString, CellStyle cellStyle) {
		boolean result = false;
		if (wb != null) {
			int sheetCount = wb.getNumberOfSheets();
			if (sheetCount > sheetIdx) {
				sheet = wb.getSheetAt(sheetIdx);
			} else {
				sheet = wb.createSheet();
				int thisSheetIdx = wb.getNumberOfSheets() - 1;
				if (sheetName != null) {
					wb.setSheetName(thisSheetIdx, sheetName);
				} else {
					wb.setSheetName(thisSheetIdx, "Sheet" + (thisSheetIdx + 1));
				}
			}
			
			
			if (sheet.getRow(rowIdx) != null) {
				row = sheet.getRow(rowIdx);
			} else {
				row = sheet.createRow(rowIdx);
			}
			if (rowHeight != 0){
				row.setHeight((short)rowHeight);
			}
			
			
			if (row.getCell(cellIdx) != null) {
				cell = row.getCell(cellIdx);
			} else {
				cell = row.createCell(cellIdx);
			}
			if (cellStyle != null) {
				cell.setCellStyle(cellStyle);
			}
			if (cellString != null) {
				cell.setCellValue(cellString);
			}
			if (cellSize !=0){
				sheet.setColumnWidth(cell.getColumnIndex(), (short)cellSize);
			}
			
			System.out.println("COMPLETE insert Image to ExcelFile = " + filePath + " insertMasage = " + cellString + "  Sheet = " + sheetIdx + "  Row = " + rowIdx + "  Cell = " + cellIdx);
			result = true;
		} else {
			result = false;
		}
		return result;
	}

	/**
	 * 엑셀 파일 내부에 그림을 넣어준다.
	 * 
	 * @param sheetName
	 *            null 입력하려고 하는 목적 시트의 이름(신규 입력시에만 변경되며 기존에 입력되어 있는 항목에 넣을 경우는
	 *            변경되지 않는다)
	 * @param sheetIdx
	 *            not null 입력하려고 하는 목적 시트의 인덱스
	 * @param rowIdx
	 *            not null 입력하려고 하는 목적 줄
	 * @param cellIdx
	 *            not null 입력하려고 하는 목적 열
	 * @param imageXsize
	 *            not null 엑셀 내부에서 이미지가 차지하는 가로 길이(엑셀 cell size 기준)
	 * @param imageYsize
	 *            not null 엑셀 내부에서 이미지가 차지하는 세로 길이(엑셀 row size 기준)
	 * @param cellStyle
	 *            null 엑셀의 cell 가질 스타일
	 * @param imageFilePath
	 *            not null 엑셀에 삽입 할 이미지가 존재하는 경로
	 * @return boolean 작업 성공여부
	 * @throws IOException
	 */
	public boolean insertImageFile(String sheetName, int sheetIdx, int rowIdx, int cellIdx,  int rowHeight,int cellSize,int imageXsize, int imageYsize, CellStyle cellStyle, String imageFilePath) throws IOException {
		boolean result = false;
		if (wb != null) {
			int sheetCount = wb.getNumberOfSheets();
			if (sheetCount > sheetIdx) {
				sheet = wb.getSheetAt(sheetIdx);
			} else {
				sheet = wb.createSheet();
				int thisSheetIdx = wb.getNumberOfSheets() - 1;
				if (sheetName != null) {
					wb.setSheetName(thisSheetIdx, sheetName);
				} else {
					wb.setSheetName(thisSheetIdx, "Sheet" + (thisSheetIdx + 1));
				}
			}
			if (sheet.getRow(rowIdx) != null) {
				row = sheet.getRow(rowIdx);
			} else {
				row = sheet.createRow(rowIdx);
			}
			if (rowHeight != 0){
				row.setHeight((short)rowHeight);
			}
			if (row.getCell(cellIdx) != null) {
				cell = row.getCell(cellIdx);
			} else {
				cell = row.createCell(cellIdx);
			}
			if (cellStyle != null) {
				cell.setCellStyle(cellStyle);
			}
			
			if (imageFilePath != null) {
				int pictureType;
				String ImagefileExtension = imageFilePath.substring(imageFilePath.lastIndexOf(".") + 1, imageFilePath.length());
				if (ImagefileExtension.toLowerCase().equals("jpg") || imageFilePath.toLowerCase().equals("jpeg")) {
					pictureType = Workbook.PICTURE_TYPE_JPEG;
				} else if (ImagefileExtension.toLowerCase().equals("png")) {
					pictureType = Workbook.PICTURE_TYPE_PNG;
				} else if (ImagefileExtension.toLowerCase().equals("pcx")) {
					pictureType = Workbook.PICTURE_TYPE_PICT;
				} else {
					pictureType = 0;
				}
				InputStream imagesStream = new FileInputStream(imageFilePath);
				byte[] bytes = IOUtils.toByteArray(imagesStream);
				if (pictureType != 0) {
					int my_picture_id = wb.addPicture(bytes, pictureType);
					imagesStream.close();
					Drawing drawing = sheet.createDrawingPatriarch();
					ClientAnchor anchor = null;
					if (fileType.equals("xls")) {
						anchor = new HSSFClientAnchor();
					} else if (fileType.equals("xlsx")) {
						anchor = new XSSFClientAnchor();
					}
					anchor.setRow1(rowIdx);
					anchor.setCol1(cellIdx);
					if (imageYsize > 0) {
						anchor.setRow2(imageYsize);
					}
					if (imageXsize > 0) {
						anchor.setCol2(imageXsize);
					}
					Picture picture = drawing.createPicture(anchor, my_picture_id);
					picture.resize();
				}
			}
			if (cellSize !=0){
				sheet.setColumnWidth(cell.getColumnIndex(), (short)cellSize);
			}
			
			for(int columnIndex = 0; columnIndex < cell.getColumnIndex()+1; columnIndex++) {
	             sheet.autoSizeColumn(columnIndex);
	        }

			
			System.out.println("COMPLETE insert Image to ExcelFile = " + filePath + " imageSourceFile = " + imageFilePath + "  Sheet = " + sheetIdx + "  Row = " + rowIdx + "  Cell = " + cellIdx);
			result = true;
		} else {
			result = false;
		}
		return result;
	}

	/**
	 * 현재 ExcelUtil의 WoorBook에 저장된 내용을 파일 쓰기 하는 함수
	 * 
	 * @return boolean 작업성공여부
	 */
	public boolean saveExcelFile() {
		boolean result = false;
		try {
			FileOutputStream fos = new FileOutputStream(filePath);
			wb.write(fos);
			fos.close();
			System.out.println("COMPLETE SAVED ExcelFile = " + filePath);
		} catch (IOException e) {
			e.printStackTrace();
			return result;
		}
		return result;
	}
}
