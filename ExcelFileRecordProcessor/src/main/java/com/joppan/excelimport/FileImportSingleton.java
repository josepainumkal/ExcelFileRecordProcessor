package com.joppan.excelimport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.atomic.AtomicInteger;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileImportSingleton {
	private static FileImportSingleton _instance;
	public ArrayBlockingQueue<Map<String,String>> queue = null;
	public AtomicInteger processedRecCnt = null;	
	
	private FileImportSingleton() {
		queue = new ArrayBlockingQueue<>(1000);
		processedRecCnt = new AtomicInteger();
	}
	
	public static FileImportSingleton getInstance() {
		if(_instance == null) {
			_instance = new FileImportSingleton();
		}
		return _instance;
	}

	public void setProcessedRecCnt(int processedRecCnt) {
		this.processedRecCnt.set(processedRecCnt);
	}

	public int getProcessedRecCnt() {
		return processedRecCnt.get();
	}
	
	public void incrementProcessedRecCnt() {
		processedRecCnt.incrementAndGet();
	}
	
	
	public synchronized void writeToOutputFile(Map<String,String> dataRowMap, String outFileName) throws Exception {
		Workbook wb = null;
		Sheet sheet = null;
		Row row = null;
		FileInputStream is = null;
		FileOutputStream os = null;
		
		try {
			File file = new File(outFileName);
			if(file.isFile()) {
				// file exist and found, then append record
				is = new FileInputStream(new File(outFileName));
				wb = WorkbookFactory.create(is);
				sheet = wb.getSheetAt(0);
				row = sheet.createRow(sheet.getLastRowNum()+1);
				writeValues(dataRowMap, row);
				is.close();
				
			}else {
				//file not found. Create the file and write record+heading
				os = FileUtils.openOutputStream(new File(outFileName));
				wb = new XSSFWorkbook();
				sheet = wb.createSheet("Out Records");
				row = sheet.createRow(0); //first line of the file created
				writeKeys(dataRowMap, row);
				row = sheet.createRow(1); //second line created
				writeValues(dataRowMap, row);
			}
			os = new FileOutputStream(outFileName);
			wb.write(os);
			wb.close();
			os.close();
			
		}catch(Exception ex) {
			ex.printStackTrace();
		}finally {
			if(is!=null) {
				is.close();
			}
			if(os!=null) {
				os.close();
			}
			if(wb!=null) {
				wb.close();
			}
		}
		
	}
	
	private void writeKeys(Map<String,String> dataRowMap, Row row) {
		Cell cell = null;
		int colCount = 0;
		List<String> keys = new ArrayList<String>(dataRowMap.keySet());
		for(String key: keys) {
			cell = row.createCell(colCount);
			cell.setCellValue(key);
			colCount++;
		}
		
	}
	
	private void writeValues(Map<String,String> dataRowMap, Row row) {
		Cell cell = null;
		int colCount = 0;
		List<String> values = new ArrayList<String>(dataRowMap.values());
		for(String value: values) {
			cell = row.createCell(colCount);
			cell.setCellValue(value);
			colCount++;
		}
		
	}
	
}
