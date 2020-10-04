package com.jitin.pojotoexcel.model;

import java.util.List;

public class ExcelRequest {
	private List<?> data;
	private String fileStoragePath;
	private String filename;
	private String sheetName;

	public ExcelRequest(String fileStoragePath, String filename, List<?> data) {
		this.fileStoragePath = fileStoragePath;
		this.filename = filename;
		this.data = data;
	}

	public List<?> getData() {
		return data;
	}

	public void setData(List<?> data) {
		this.data = data;
	}

	public String getFileStoragePath() {
		return fileStoragePath;
	}

	public void setFileStoragePath(String fileStoragePath) {
		this.fileStoragePath = fileStoragePath;
	}

	public String getFilename() {
		return filename;
	}

	public void setFilename(String filename) {
		this.filename = filename;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

}
