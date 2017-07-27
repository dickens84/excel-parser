package com.dickens.core.parser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * It is a xls file reader implemented using Iterator design pattern.
 * It internally uses default POI APIs for reading xls files.
 * 
 ***********************************************************
 ################## Recommended Use: ########################
 #  XLSReader reader = new XLSReader(filePath);           #
 #  Iterator<List<String>> iterator = reader.getIterator(); #
 #  while(iterator.hasNext()){                              #
 #	List<String> row = iterator.next();                     #
 #	if(row!=null){                                          #
 #		 for(String data:row){                              #
 #			// process data	here                            #
 #	     }                                                  #
 #   }                                                      #
 # reader.close();                                          #
 ############################################################ 
 * 
 *  			
 * @author Dickens Prabhu
 *
 */
public class XLSReader extends GenericFileReader{

	/** The logger. */
	private static Logger logger = Logger.getLogger(XLSReader.class.getName());
	
	/** The i stream. */
	private InputStream iStream;
	
	/** The wb. */
	private Workbook wb;
	
	/** The sheet. */
	private Sheet sheet;
	
	/** The read empty row. */
	//true = empty rows will be read
	private final boolean readEmptyRow;

	/**
	 * Performs Setup, Empty rows will also be read.
	 *
	 * @author Dickens Prabhu
	 * @param filePath the file path
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public XLSReader(String filePath) throws IOException{
		this(filePath,true);
	}

	/**
	 * Take input stream as input and perform setup (Empty rows will also be read).
	 * If input stream is null it will throw FileNotFoundException.
	 *
	 * @author Dickens Prabhu
	 * @param iStream the i stream
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public XLSReader(InputStream iStream) throws IOException{
		this(iStream,true);
	}

	
	/**
	 * Performs Setup.
	 * If file does not exist then FileNotFoundException will be thrown.
	 *
	 * @author Dickens Prabhu
	 * @param filePath (path of the .xls input file)
	 * @param readEmptyRow (if false then empty rows will not be read)
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public XLSReader(String filePath,boolean readEmptyRow) throws IOException{
		File xlsxFile = new File(filePath);
		if (!xlsxFile.exists()) {
			logger.info("Not found or not a file: " + xlsxFile.getPath());
			throw new FileNotFoundException("Not found or not a file: " + xlsxFile.getPath());
		}
		this.readEmptyRow=readEmptyRow;
		this.iStream = new FileInputStream(xlsxFile);
		process(this.iStream);
	}

	/**
	 * Take input stream as input and perform setup.
	 * If input stream is null it will throw FileNotFoundException.
	 *
	 * @author Dickens Prabhu
	 * @param iStream (input stream for the .xls file)
	 * @param readEmptyRow the read empty row
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public XLSReader(InputStream iStream,boolean readEmptyRow) throws IOException{
		if (iStream==null) {
			logger.info("Input Stream is Null");
			throw new FileNotFoundException("Input Stream is Null");
		}
		this.readEmptyRow=readEmptyRow;
		process(iStream);
	}
	
	
	/**
	 * Return iterator for reading .xls file
	 *
	 * @author Dickens Prabhu
	 * @return the iterator
	 */
	@Override
	public Iterator<List<String>> getIterator() {
		return new XLSIterator();
	}

	/**
	 * Perform resource cleanup like closing opened streams,etc.
	 *
	 * @author Dickens Prabhu
	 * @throws Exception the exception
	 */
	@Override
	public void close() throws Exception {
		if(iStream!=null){
			iStream.close();
		}
	}

	/**
	 * Initiates the processing of the XLSX file.
	 * The current implementation process the first sheet only.
	 *
	 * @param iStream the i stream
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	private void process(InputStream iStream) throws IOException{
		wb = new HSSFWorkbook(iStream);
		sheet = wb.getSheetAt(wb.getActiveSheetIndex()); // get active sheet
		
	}
	
	/**
	 * Provide implementation of the Iterator interface for iterating 
	 * over the rows of input xlsx file.
	 * 
	 * @author Dickens Prabhu
	 *
	 */
	private class XLSIterator implements Iterator<List<String>>{

		/** points to the current column being referenced. */
		private int thisColumn = -1;
		
		/** The last column printed to the output stream */
		private int lastColumnNumber = -1;

		/** holds true for the first row(header row), else holds false; */
		private boolean isHeader=true;
		
		/** hold the number of headerColumns, used to generate empty string for last empty columns */
		private int numberOfHeaders = 0;
		
		/** private Row currentRow; */
		private List<String> currentRow;
		
		/** row iterator from poi */
		private Iterator<Row> rowIterator;
		
		/**
		 * Performs initialization.
		 */
		public XLSIterator() {
			this.currentRow = new ArrayList<String>();
			rowIterator = sheet.iterator();
		}
		
		
		/**
		 * Uses default row iterator hasNext() method implementation.
		 *
		 * @author Dickens Prabhu
		 * @return true, if successful
		 */
		public boolean hasNext() {
			return hasNextRow();
		}

		/**
		 * if readEmptyRow is set to false then this method will look for the next non empty row and will return true after updating the current row.
		 * otherwise it check for the next available row , update current row & return.
		 *
		 * @author Dickens Prabhu
		 * @return true, if successful
		 */
		private boolean hasNextRow(){
			if(readEmptyRow){
				if(rowIterator.hasNext()){
					nextRow();
					return true;
				}else{
					return false;
				}
			}else{
				while(rowIterator.hasNext()){
					nextRow();
					if(!isEmptyCurrentRow(this.currentRow)){
						return true;
					}
				}
			}
			return false;
		}
		
		/**
		 * This method returns list of cell data of current row.
		 * This method should not be called before calling hasNext() method.
		 * Empty string("") is returned as value of empty cell.
		 *
		 * @author Dickens Prabhu
		 * @return the list
		 */
		public List<String> next() {
			return this.currentRow;
		}

		/**
		 * This method returns list of cell data of next row.
		 * This method should not be called before calling hasNext() method.
		 * Empty string("") is returned as value of empty cell.
		 *
		 * @author Dickens Prabhu
		 * @return the list
		 */
		@SuppressWarnings("deprecation")
		private List<String> nextRow() {
			String thisStr = null;
			int cellType;
			clearCurrentRow();
			lastColumnNumber=-1;
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while(cellIterator.hasNext()){
				Cell cell = cellIterator.next();
				thisColumn= cell.getColumnIndex();
				cellType = cell.getCellType();
				CellStyle style = cell.getCellStyle();
				int formatIndex=-1;
				String formatString=null;
				if(style!=null){
					formatIndex = style.getDataFormat();
					formatString = style.getDataFormatString();
					if (formatString == null){
						formatString = BuiltinFormats.getBuiltinFormat(formatIndex);	
					}
				}				
				
				if(cellType == Cell.CELL_TYPE_BOOLEAN){
					if(cell.getBooleanCellValue()){
						thisStr="TRUE";
					}else{
						thisStr="FALSE";
					}
				}
				else if(cellType==Cell.CELL_TYPE_NUMERIC){

					if(HSSFDateUtil.isCellDateFormatted(cell)){
						thisStr = cell.toString();
						
					}else{
						thisStr = cell.toString(); // unformatted numeric value

					}
				}
				else if(cellType==Cell.CELL_TYPE_FORMULA){
					 switch(cell.getCachedFormulaResultType()) {
			            case Cell.CELL_TYPE_NUMERIC:
							if(HSSFDateUtil.isCellDateFormatted(cell)){
								thisStr = cell.toString();
							}else{
								thisStr = String.valueOf(cell.getNumericCellValue()); // unformatted numeric value

							}

			               break;
			            case Cell.CELL_TYPE_STRING:
			                thisStr = cell.getRichStringCellValue().toString();
			                break;
			        }
				}
				else if(cellType==Cell.CELL_TYPE_ERROR){
					thisStr = cell.toString();					
				}
				else if(cellType==Cell.CELL_TYPE_BLANK){
					thisStr = "";
				}else if(cellType==Cell.CELL_TYPE_STRING){
					thisStr = cell.getStringCellValue();
				}
				
				// counting the number of header columns. Header columns cannot be empty hence empty checking logic is not required.
				if(isHeader){
					++numberOfHeaders;
				}
				
	          
	             for (int i = lastColumnNumber; i < thisColumn-1; i++){
	                 updateCurrentRow("");
	             }

				updateCurrentRow(thisStr);
	            
				// Update column
	             if (thisColumn > -1){
	                 lastColumnNumber = thisColumn;
	             }
			}
            // Print out any missing commas if needed for rows other than header row.
            if (!isHeader && numberOfHeaders > 0) {
               
                for (int i = lastColumnNumber; i < numberOfHeaders-1; i++) {
               	 updateCurrentRow("");
                }
            }
			
			if(isHeader){
				isHeader=false;
			}
			
			return this.currentRow;
		}
		
		
		
		/**
		 * NOT SUPPORTED IN THE CURRENT IMPLEMENTATION.
		 * @author Dickens Prabhu
		 */
		public void remove() {
			throw new UnsupportedOperationException();
		}
		
		
		/**
		 * Reset the currentRow by removing its content.
		 * @author Dickens Prabhu
		 */
		private void clearCurrentRow(){
			if(this.currentRow!=null){
				this.currentRow.clear();
			}
		}
		
		/**
		 * Update the current row by adding cell data to it.
		 *
		 * @author Dickens Prabhu
		 * @param cellData the cell data
		 */
		private void updateCurrentRow(String cellData){
			if(this.currentRow!=null){
				this.currentRow.add(cellData);
			}
		}
		
	}

}
