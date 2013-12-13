package com.soc.excel;

import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelArray{
	// excel row 좌표값
	public int excelRowIndex;
	// excel cell 좌표값
	public int excelCulIndex;
	// excel sheet 페이지
	public int excelSheet;
	// excel sheet 이름
	public String SheetName;
	// cell 에 입력할 텍스트 내용
	public String textString;
	// image를 넣을 경우 이미지 경로
	public String imagePath;
	// image의 x사이즈(cell 크기로 지정)
	public int imageXSize = 1;
	// image의 y사이즈(row 크기로 지정)
	public int imageYSize = 1;
	// 경우 엑셀 스타일 지정
	public CellStyle cellStyle;
	
	
	public void setCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}

	public CellStyle getCellStyle() {
		return cellStyle;
	}

	public void setImageXSize(int imageXSize) {
		this.imageXSize = imageXSize;
	}

	public int getImageXSize() {
		return imageXSize;
	}

	public void setImageYSize(int imageYSize) {
		this.imageYSize = imageYSize;
	}

	public int getImageYSize() {
		return imageYSize;
	}

	public void setSheetName(String sheetName) {
		SheetName = sheetName;
	}

	public String getSheetName() {
		return SheetName;
	}

	public void setImagePath(String imagePath) {
		this.imagePath = imagePath;
	}

	public String getImagePath() {
		return imagePath;
	}

	public void setTextString(String excelText) {
		this.textString = excelText;
	}

	public String getTextString() {
		return textString;
	}

	public void setExcelCulIndex(int excelCulIndex) {
		this.excelCulIndex = excelCulIndex;
	}

	public int getExcelCulIndex() {
		return excelCulIndex;
	}

	public void setExcelRowIndex(int excelRowIndex) {
		this.excelRowIndex = excelRowIndex;
	}

	public int getExcelRowIndex() {
		return excelRowIndex;
	}

	public void setExcelSheet(int excelSheet) {
		this.excelSheet = excelSheet;
	}

	public int getExcelSheet() {
		return excelSheet;
	}
}
