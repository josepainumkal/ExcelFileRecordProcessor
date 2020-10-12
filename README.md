# ExcelFileRecordProcessor
Multi-threaded java solution to process records in an excel file using producer-consumer problem. 

Accomplished below in this project:
1. Implemented a producer-consumer multi-threaded solution to process the data in an excel file. 
   - RecordProducer read records from excel file and put the record in a queue.
   - RecordConsumer read the record from the queue and process the record (in this example, write the record into a another excel file)
2. Reading from an excel file (using apache poi)
3. Writing to an excel file (using apache poi)
4. Implemented a java gradle project

## To run the project:
```bash
Run the file RecordProducer.java (contains the main())
```

## Jars Required:
```bash
        'org.apache.poi:poi:4.1.2',
        'org.apache.poi:poi-ooxml:4.1.2',
        'org.apache.poi:poi-ooxml-schemas:4.1.2',
        'org.apache.xmlbeans:xmlbeans:3.1.0',      
        'org.apache.commons:commons-collections4:4.4',
        'org.apache.commons:commons-compress:1.20',
        'commons-io:commons-io:2.5'
```

## Useful tips:
To upload/save an excel file to BLOB field of a table in Oracle DB:

```bash
String importFile = "C:/folder/filename.xlsx";
Connection conn = getOracleConnection();
String insertQuery = "INSERT INTO <Tablename>(column Name) VALUES(?)";
PreparedStatement ps = conn.prepareStatement(insertQuery);
byte[] fileContent = FileUtils.readFileToByteArray(new File(importFile));
InputStream in = new ByteArrayInputStream(fileContent);
ps.setBinaryStream(1,in);
ps.executeQuery();
```
To read excel file saved in blob field in Oracle DB:
```bash
String query = "SELECT * FROM TableName";
PreparedStatement ps = conn.prepareStatement(query);
ResultSet rs = ps.executeQuery();
while(rs.next()){
   Workbook wb = new XSSFWorkbook(rs.getBinaryStream(1)); // the no of the blob field in the select query used
   String importFile = "<folderpath>/filename.xlsx";
   FileOutputStream os = FileUtils.openOutputStream(new File(importFile));
   wb.write(os);
   wb.close();
   os.close();
}
```











