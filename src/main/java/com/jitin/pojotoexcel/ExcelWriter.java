package com.jitin.pojotoexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Objects;

import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.jitin.pojotoexcel.annotations.ExcelColumn;
import com.jitin.pojotoexcel.annotations.ExcelIgnore;
import com.jitin.pojotoexcel.exception.ExcelWriterException;
import com.jitin.pojotoexcel.model.ExcelRequest;
import com.jitin.pojotoexcel.util.Constants;

public class ExcelWriter {
	private final static Logger LOG = LogManager.getLogger(ExcelWriter.class);

	public static String write(ExcelRequest request) {
		BasicConfigurator.configure(); // --Log4j configuration.
		if (Objects.isNull(request) || Objects.isNull(request.getData()) || request.getData().size() < 1) {
			LOG.error("Request object or data was null.");
			return Constants.STATUS_FAILURE;
		}
		try {
			if (new File(request.getFileStoragePath() + "/" + request.getFilename()).exists()) {
				return updateExistingExcelFile(request) ? Constants.STATUS_SUCCESS : Constants.STATUS_FAILURE;
			} else {
				return writeNewExcelFile(request) ? Constants.STATUS_SUCCESS : Constants.STATUS_FAILURE;
			}
		} catch (ExcelWriterException e) {
			LOG.error(e);
			return Constants.STATUS_FAILURE;
		}
	}
	
	private static boolean updateExistingExcelFile(ExcelRequest request) {
		String filename = new StringBuilder(request.getFileStoragePath()).append("/").append(request.getFilename()).toString();
		try(Workbook workbook =new XSSFWorkbook(new FileInputStream(new File(filename)))){
			if(writeData(workbook.getSheetAt(0),request)) {
				writeWorkbookDataToFile(workbook,request);
				return true;
			}
			return false;
		}catch(Exception e) {
			throw new ExcelWriterException(e);
		}
	}

	private static boolean writeNewExcelFile(ExcelRequest request) {
		try(Workbook workbook = new XSSFWorkbook()){
			if(StringUtils.isBlank(request.getSheetName())) {
				request.setSheetName(Constants.DEFAULT_SHEET_NAME);
			}
			Sheet sheet = workbook.createSheet(request.getSheetName());
			writeHeader(sheet, request.getData().get(0));
			if(writeData(sheet,request)) {
				writeWorkbookDataToFile(workbook,request);
				return true;
			}
			return false;
		}catch(Exception e) {
			throw new ExcelWriterException(e);
		}
	}
	
	private static boolean writeData(Sheet sheet, ExcelRequest request) {
		int rowNumber = sheet.getLastRowNum();
		for(Object obj : request.getData()) {
			Row row = sheet.createRow(++rowNumber);
			Field[] fields = obj.getClass().getDeclaredFields();
			int cellNumber=0;
			for(Field field: fields) {
				if(!field.isAnnotationPresent(ExcelIgnore.class)) {
					field.setAccessible(true);
					try {
						Object value = field.get(obj);
						Cell cell = row.createCell(cellNumber++);
						if(Objects.nonNull(value)) {
							if(value instanceof Date) {
								try {
									cell.setCellValue(new SimpleDateFormat(Constants.DEFAULT_DATE_FORMAT).format((Date)value));
								}catch(Exception e) {
									cell.setCellValue(value.toString());
								}
							}else {
								cell.setCellValue(value.toString());
							}
						}else {
							cell.setCellValue("");
						}
					}catch(Exception e) {
						LOG.error("Error occurred while reading value of "+field.getName()+" property.",e);
						throw new ExcelWriterException("Error occurred while reading value of "+field.getName()+" property.",e);
					}
				}
			}
		}
		return true;
	}

	private static void writeWorkbookDataToFile(Workbook workbook, ExcelRequest request) {
		try(OutputStream os =new FileOutputStream(new File(new StringBuilder(request.getFileStoragePath()).append("/").append(request.getFilename()).toString()))){
			workbook.write(os);
		}catch(Exception e) {
			throw new ExcelWriterException(e);
		}
	}
	
	private static void writeHeader(Sheet sheet, Object obj) {
		int rowNumber = sheet.getLastRowNum();
		Field[] fields = obj.getClass().getDeclaredFields();
		Row row = sheet.createRow(++rowNumber);
		int cellNumber = 0;
		for(int i=0;i<fields.length;i++) {
			if(fields[i].isAnnotationPresent(ExcelColumn.class)) {
				Cell cell = row.createCell(cellNumber++);
				cell.setCellValue(fields[i].getAnnotation(ExcelColumn.class).name());
			}else if(!fields[i].isAnnotationPresent(ExcelIgnore.class)) {
				Cell cell = row.createCell(cellNumber++);
				cell.setCellValue(fields[i].getName());
			}
		}
		
	}
}
