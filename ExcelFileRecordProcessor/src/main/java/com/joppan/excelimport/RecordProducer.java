package com.joppan.excelimport;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class RecordProducer {
	
	/**
	 * Read the header row in the excel file and populate headers/column position index in headerPosNameMap
	 * for example, headerPosNameMap: {0=Name, 1=Age, 2=Strike Rate, 3=Status, 4=Category}
	 * @param headerPosNameMap
	 * @param importFile
	 */
	
	private static void readHeaderLine(Map<Integer,String> headerPosNameMap, String importFile) {
		//read headerline in the excel file
		Workbook wb = null;
		FileInputStream is = null;
		DataFormatter df = null;
		String cellValue = null;
		Row row = null;
		
		try {
			is = new FileInputStream(new File(importFile));
			df = new DataFormatter();
			wb = WorkbookFactory.create(is);
			row = (Row)wb.getSheetAt(0).getRow(0); //header line
			
			if(row!=null) {
				for(int colIndex = 0; colIndex<row.getLastCellNum(); colIndex++) {
					cellValue = df.formatCellValue(row.getCell(colIndex)).trim();
					headerPosNameMap.put(colIndex, cellValue);				
				}
			}else {
				throw new Exception ("File should have header fields in the first row. Please correct.");
			}
		}
		catch(Exception ex) {
			ex.printStackTrace();
		}
	}
	
	/**
	 * Read the records in the excel file one by one and populate the record in dataRowMap
	 * and put dataRowMap in the queue
	 * @param headerPosNameMap
	 * @param importFile
	 * @param recProcessorThreadCnt
	 */
	private static void putRecordsInQueue(Map<Integer,String> headerPosNameMap,
			String importFile, int recProcessorThreadCnt) {
		//read records in the excel file
		Workbook wb = null;
		FileInputStream is = null;
		DataFormatter df = null;
		Sheet sheet = null;
		String cellValue = null;
		Row row = null;
		
		try {
			is = new FileInputStream(new File(importFile));
			df = new DataFormatter();
			wb = WorkbookFactory.create(is);
			sheet = wb.getSheetAt(0);
			
			//read from row 1 to exclude the header line
			for(int rowIndex = 1; rowIndex<=sheet.getLastRowNum(); rowIndex++) {
				row = (Row)sheet.getRow(rowIndex);
				if(row!=null) {
					Map<String, String> dataRowMap = new HashMap<String,String>();
					for(int colIndex=0; colIndex<headerPosNameMap.size(); colIndex++) {
						cellValue = df.formatCellValue(row.getCell(colIndex)).trim();
						dataRowMap.put(headerPosNameMap.get(colIndex), cellValue);
					}
					FileImportSingleton.getInstance().queue.put(dataRowMap);
				} 
			}
			
			// add poison pills to the queue to terminate consumer threads
			Map<String,String> poisonPill = new HashMap<String,String>();
			poisonPill.put("END_REC_CONSUMER", null);
			for(int i = 1; i<=recProcessorThreadCnt; i++) {
				while(!FileImportSingleton.getInstance().queue.offer(poisonPill)) {
					try {
						TimeUnit.SECONDS.sleep(1);
					}catch(InterruptedException ex) {						
					}
				}
			}
			System.out.println("Poison pill records submitted...");
			
		}catch(Exception ex) {
			ex.printStackTrace();
		}
		
	}
	
	
	
	public static void main(String args[]) {
		
		String importFile = "D:/test_files/cricPlayers.xlsx";
		String outFileName = "D:/test_files/outFile.xlsx";
		Map<Integer,String> headerPosNameMap = new HashMap<Integer,String>();
		int recProcessorThreadCnt = 4; //no of consumer threads
		System.out.println("No. of consumer threads: "+recProcessorThreadCnt);

		try {
			readHeaderLine(headerPosNameMap, importFile);
			System.out.println("headerPosNameMap: "+headerPosNameMap);	
			
			FileImportSingleton.getInstance().setProcessedRecCnt(0);
			RecordConsumer recConsumer = new RecordConsumer(outFileName);
			
			/**
			 * Start consumer threads to read from queue.
			 */
			ThreadPoolExecutor executor = (ThreadPoolExecutor) Executors.newFixedThreadPool(recProcessorThreadCnt);
			for(int i=1; i<=recProcessorThreadCnt; i++) {
				executor.submit(recConsumer);
			}

			putRecordsInQueue(headerPosNameMap,importFile,recProcessorThreadCnt);
			
			/**
			 * Shutdown consumer threads and wait for consumers to terminate
			 */
			executor.shutdown();
			while(!executor.isTerminated()) {
				try {
					executor.awaitTermination(24, TimeUnit.HOURS);
				}catch(InterruptedException ex) {
					ex.printStackTrace();
				}
			}
			System.out.println("Consumer Threads Terminated");
			
			int processedRecCount = FileImportSingleton.getInstance().getProcessedRecCnt();
			System.out.println("No. of records processed by consumer threads: "+processedRecCount);
			
			
		}catch(Exception ex) {
			ex.printStackTrace();
		}
		
	}

}
