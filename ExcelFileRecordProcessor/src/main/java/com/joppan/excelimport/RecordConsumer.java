package com.joppan.excelimport;
import java.util.Map;

public class RecordConsumer implements Runnable{
	private String outFileName;
	
	public RecordConsumer(String outFileName) {
		this.outFileName = outFileName;
	}

	@Override
	public void run() {
		System.out.println("Starting thread to consumer record");
		boolean isActive = true;
		while(isActive) {
			Map<String,String> dataRowMap = FileImportSingleton.getInstance().queue.poll();
			if(dataRowMap!=null) {
				if(dataRowMap.containsKey("END_REC_CONSUMER")) {
					//poison pill detected
					System.out.println("Poison pill recieved in thread \'"+Thread.currentThread().getName()+"\'..Exiting the loop");
					isActive = false;
					break;					
				}else {
					try {
						System.out.println("Record received in thread \'"+Thread.currentThread().getName()+
								   "\' dataRowMap: "+dataRowMap);
						FileImportSingleton.getInstance().incrementProcessedRecCnt();
						
						//add some logic here to process the record...Here, I am writing the record to an excel file
						FileImportSingleton.getInstance().writeToOutputFile(dataRowMap, outFileName);
					}catch(Exception ex) {
						ex.printStackTrace();
					}
				}		
			}			
		} //end while loop
	}
}
