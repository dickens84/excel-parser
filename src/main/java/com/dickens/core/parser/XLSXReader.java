package com.dickens.core.parser;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Logger;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.SAXException;

/**
 * This is a fast xlsx file reader.
 * It internally uses STAX parser to process the xlsx file.
 * 
 ***********************************************************
 ################## Recommended Use: ########################
 #  XLSXReader reader = new XLSXReader(filePath);           #
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
public class XLSXReader extends GenericFileReader{
	
	private static Logger logger = Logger.getLogger(XLSXReader.class.getName());
	/**
	 * The type of the data value is indicated by an attribute on the cell.
	 * The value is usually in a "v" element within the cell.
	 */
	enum xssfDataType {
		BOOL,
		ERROR,
		FORMULA,
		INLINESTR,
		SSTINDEX,
		NUMBER,
	}

	private final XMLInputFactory factory;
	private XMLStreamReader streamReader;
	private OPCPackage opcPackage;
	
	
	/**
	 * Table with styles
	 */
	private StylesTable stylesTable;

	/**
	 * Table with unique strings
	 */
	private ReadOnlySharedStringsTable sharedStringsTable;

	//true = empty rows will be read
	private final boolean readEmptyRow;
	
	/**
	 * Take xlsx file path and perform setup.
	 * If file doesnot exist it will throw FileNotFoundException
	 * 
	 * @author Dickens Prabhu
	 * @param filePath
	 * @throws IOException
	 * @throws OpenXML4JException
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 * @throws XMLStreamException
	 */
	
	public XLSXReader(String filePath) throws IOException, OpenXML4JException, ParserConfigurationException, SAXException, XMLStreamException {
		this(filePath,true);
	}

	/**
	 * Take input stream as input and perform setup.
	 * If input stream is null it will throw FileNotFoundException
	 * 
	 * @author Dickens Prabhu
	 * @param iStream
	 * @throws IOException
	 * @throws OpenXML4JException
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 * @throws XMLStreamException
	 */
	public XLSXReader(InputStream iStream) throws IOException, OpenXML4JException, ParserConfigurationException, SAXException, XMLStreamException{
		this(iStream,true);
	}
	
	
	/**
	 * Take xlsx file path and perform setup.
	 * If file doesnot exist it will throw FileNotFoundException
	 * 
	 * @author Dickens Prabhu
	 * @param filePath
	 * @throws IOException
	 * @throws OpenXML4JException
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 * @throws XMLStreamException
	 */
	
	public XLSXReader(String filePath,boolean readEmptyRow) throws IOException, OpenXML4JException, ParserConfigurationException, SAXException, XMLStreamException {
		File xlsxFile = new File(filePath);
		if (!xlsxFile.exists()) {
			logger.info("Not found or not a file: " + xlsxFile.getPath());
			throw new FileNotFoundException("Not found or not a file: " + xlsxFile.getPath());
		}
		this.factory = XMLInputFactory.newInstance();
		this.readEmptyRow=readEmptyRow;
		// The package open is instantaneous, as it should be.
		opcPackage = OPCPackage.open(xlsxFile.getPath(), PackageAccess.READ);
		process(opcPackage);
	}

	/**
	 * Take input stream as input and perform setup.
	 * If input stream is null it will throw FileNotFoundException
	 * 
	 * @author Dickens Prabhu
	 * @param iStream
	 * @throws IOException
	 * @throws OpenXML4JException
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 * @throws XMLStreamException
	 */
	public XLSXReader(InputStream iStream,boolean readEmptyRow) throws IOException, OpenXML4JException, ParserConfigurationException, SAXException, XMLStreamException{
		if (iStream==null) {
			logger.info("Input Stream is Null");
			throw new FileNotFoundException("Input Stream is Null");
		}
		this.factory = XMLInputFactory.newInstance();
		this.readEmptyRow=readEmptyRow;
		// The package open is instantaneous, as it should be.
		opcPackage = OPCPackage.open(iStream);
		process(opcPackage);
	}
	
	
	/**
	 * Returns the iterator for reading xlsx files.
	 * @author Dickens Prabhu
	 * @return
	 */
	@Override
	public Iterator<List<String>> getIterator(){
		return new XLSXIterator();
	}
	
	/**
	 * Perform resource cleanup like closing opened streams,etc.
	 * @author Dickens Prabhu
	 * @throws IOException
	 */
	@Override
	public void close() throws Exception{
		if(opcPackage!=null){
			opcPackage.close();
		}
		if(streamReader!=null){
			streamReader.close();
		}
	}
	
	/**
	 * Initiates the processing of the XLSX file.
	 * The current implementation process the Active Sheet.
	 * 
	 * @author Dickens Prabhu
	 * @throws IOException
	 * @throws OpenXML4JException
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 * @throws XMLStreamException 
	 */
	private void process(OPCPackage opcPackage) throws IOException, OpenXML4JException, ParserConfigurationException, SAXException, XMLStreamException {

		this.sharedStringsTable = new ReadOnlySharedStringsTable(opcPackage);
		XSSFReader xssfReader = new XSSFReader(opcPackage);
		this.stylesTable = xssfReader.getStylesTable();
		
		XSSFReader.SheetIterator dataItr = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
		XSSFReader.SheetIterator sheetItr = (XSSFReader.SheetIterator) xssfReader.getSheetsData(); // extra iterator for identifying active sheet
		boolean isProcessed=false;
		while (dataItr.hasNext() && sheetItr.hasNext()) {
			InputStream dataStream = dataItr.next();
			InputStream sheetStream = sheetItr.next(); // Not able to re use input stream thus getting extra input stream for identifying active sheet.
			
			if(isActiveSheet(sheetStream)){   
				processSheet(dataStream);
				isProcessed=true;
				break; // Process only the active sheet
			}
		}
		if (!isProcessed) { // it means there is no active tab, select the first sheet.
			while (dataItr.hasNext() && sheetItr.hasNext()) {
				InputStream dataStream = dataItr.next();
				processSheet(dataStream);
				break; // Process only the active sheet

			}
		}

	}
	
	/**
	 * This method checks for the value of "tabSelected" attribute of "SheetView" XML
	 * element to find out whether the sheet is active or not.
	 * "SheetView" XML element contains attribute "tabSelected" whose value is "1" for 
	 * active sheet.
	 * 
	 * @author Dickens Prabhu
	 * @param sheetInputStream
	 * @return
	 * @throws XMLStreamException
	 */
	private boolean isActiveSheet(InputStream sheetInputStream) throws XMLStreamException{
		XMLStreamReader xmlStreamReader = this.factory.createXMLStreamReader(sheetInputStream);
		while(xmlStreamReader.hasNext()){
			xmlStreamReader.next();
			if(xmlStreamReader.getEventType() == XMLStreamReader.START_ELEMENT){
				String name=xmlStreamReader.getLocalName();
				if("sheetView".equals(name)){
					String tabSelected = xmlStreamReader.getAttributeValue(null, "tabSelected");
					if(tabSelected!=null && "1".equals(tabSelected)){
						return true;
					}
					break;// break once sheetView is found
				}
			}
		}
		return false;
	}
	
	
	/**
	 * Parses and shows the content of one sheet
	 * using the specified styles and shared-strings tables.
	 * @author Dickens Prabhu
	 * @param styles
	 * @param strings
	 * @param sheetInputStream
	 * @throws XMLStreamException 
	 */
	private void processSheet(InputStream sheetInputStream)	throws IOException, ParserConfigurationException, SAXException, XMLStreamException {
		this.streamReader = this.factory.createXMLStreamReader(sheetInputStream);
	}

	/**
	 * Provide implementation of the Iterator interface for iterating 
	 * over the rows of input xlsx file.
	 * 
	 * @author Dickens Prabhu
	 *
	 */
	private class XLSXIterator implements Iterator<List<String>>{
		
		// Set when V start element is seen
		private boolean vIsOpen;

		// Set when cell start element is seen;
		// used when cell close element is seen.
		private xssfDataType nextDataType;

		// Used to format numeric cell values.
		private short formatIndex;
		private String formatString;
		private final DataFormatter formatter;

		// points to the current column being referenced.
		private int thisColumn = -1;
		// The last column printed to the output stream
		private int lastColumnNumber = -1;

		// holds true for the first row(header row), else holds false;
		private boolean isHeader=true;
		
		// hold the number of headerColumns, used to generate empty string for last empty columns
		private int numberOfHeaders = 0; 
		
		// Gathers characters as they are seen.
		private StringBuffer value;
		
		//private Row currentRow;
		private List<String> currentRow;

		/**
		 * Performs initialization
		 */
		public XLSXIterator(){
			this.currentRow = new ArrayList<String>();
			this.value = new StringBuffer();
			this.nextDataType = xssfDataType.NUMBER;
			this.formatter = new DataFormatter() {
				 public String formatRawCellContents(double value, int formatIndex, String formatString, boolean use1904Windowing) {
					 if(DateUtil.isADateFormat(formatIndex,formatString)) {
						 formatString="MM/dd/yyyy";
					 }
					return  super.formatRawCellContents(value, formatIndex, formatString,use1904Windowing);
				 }
			};
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
		 * @author Dickens Prabhu
		 * @param cellData
		 */
		private void updateCurrentRow(String cellData){
			if(this.currentRow!=null){
				this.currentRow.add(cellData);
			}
		}
		
		
		/**
		 * This method checks if next row is present or not.
		 * if readEmptyRow is set to false then only non empty rows are considered.
		 * 
		 * @author Dickens Prabhu
		 */
		public boolean hasNext() {
			try{
				if(readEmptyRow){
					if(hasNextRow()){ // if nextRow is present
						nextRow(); // move pointer to the next row
						return true;
					}else{
						return false;
					}
				}else{
					while(hasNextRow()){
						nextRow();
						if(!isEmptyCurrentRow(this.currentRow)){
							return true;
						}
					}	
				}
			} catch (XMLStreamException e) {
				e.printStackTrace();
			}
			return false;
		}
		
		/**
		 * This method loop through the xlm elements to find the row element.
		 * If row element is found then it returns true else false.
		 * @author Dickens Prabhu
		 * @return true if going forward row element is present, else false
		 * @throws XMLStreamException
		 */
		private boolean hasNextRow() throws XMLStreamException{
			boolean hasNextRow=false;
			while(!hasNextRow && streamReader.hasNext()){
				streamReader.next();
				if(streamReader.getEventType() == XMLStreamReader.START_ELEMENT){
					String name=streamReader.getLocalName();
					if("row".equals(name)){
						clearCurrentRow();
						hasNextRow=true;
					}
				}
			}
			return hasNextRow;
		}

		/**
		 * This method returns list of cell data of current row.
		 * This method should not be called before calling hasNext() method.
		 * Empty string("") is returned as value of empty cell.
		 * @author Dickens Prabhu
		 */
		public List<String> next() {
			return this.currentRow;
		}

		
		/**
		 * This method returns list of cell data of next row.
		 * This method should not be called before calling hasNext() method.
		 * Empty string("") is returned as value of empty cell.
		 * @author Dickens Prabhu
		 */
		private List<String> nextRow() {
			boolean currentRowEnds = false;
			try {
					while(!currentRowEnds){
						// xml element start
						if(streamReader.getEventType() == XMLStreamReader.START_ELEMENT){
							String name=streamReader.getLocalName();
							if ("inlineStr".equals(name) || "v".equals(name)) {
								vIsOpen = true;
								// Clear contents cache
								value.setLength(0);
							}
							// c => cell
							else if ("c".equals(name)) {
								// Get the cell reference
								String r = streamReader.getAttributeValue(null,"r");
								int firstDigit = -1;
								for (int c = 0; c < r.length(); ++c) {
									if (Character.isDigit(r.charAt(c))) {
										firstDigit = c;
										break;
									}
								}
								thisColumn = nameToColumn(r.substring(0, firstDigit));
		
								// Set up defaults.
								this.nextDataType = xssfDataType.NUMBER;
								this.formatIndex = -1;
								this.formatString = null;
								String cellType = streamReader.getAttributeValue(null,"t");
								String cellStyleStr = streamReader.getAttributeValue(null,"s");
								if ("b".equals(cellType))
									nextDataType = xssfDataType.BOOL;
								else if ("e".equals(cellType))
									nextDataType = xssfDataType.ERROR;
								else if ("inlineStr".equals(cellType))
									nextDataType = xssfDataType.INLINESTR;
								else if ("s".equals(cellType))
									nextDataType = xssfDataType.SSTINDEX;
								else if ("str".equals(cellType))
									nextDataType = xssfDataType.FORMULA;
								else if (cellStyleStr != null) {
									// It's a number, but almost certainly one
									//  with a special style or format 
									XSSFCellStyle style = null;
									if (cellStyleStr != null) {
										int styleIndex = Integer.parseInt(cellStyleStr);
										style = stylesTable.getStyleAt(styleIndex);
									}
									if (style == null && stylesTable.getNumCellStyles() > 0) {
										style = stylesTable.getStyleAt(0);
									}
									if (style != null) {
										this.formatIndex = style.getDataFormat();
										this.formatString = style.getDataFormatString();
										if (this.formatString == null)
											this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
									}
								}
							}
		
					    }else if(streamReader.getEventType() == XMLStreamReader.CHARACTERS){
					    	if (vIsOpen){
								value.append(streamReader.getText());
							}
					    }else if(streamReader.getEventType() == XMLStreamReader.END_ELEMENT){
					    	
							String thisStr = null;
							String name=streamReader.getLocalName();
							// v => contents of a cell
							if ("v".equals(name)) {
								// Process the value contents as required.
								// Do now, as characters() may be called more than once
								switch (nextDataType) {
		
								case BOOL:
									char first = value.charAt(0);
									thisStr = first == '0' ? "FALSE" : "TRUE";
									break;
		
								case ERROR:
									thisStr = value.toString();
									break;
		
								case FORMULA:
									// A formula could result in a string value,
									// so always add double-quote characters.
									thisStr = '"' + value.toString() + '"';
									break;
		
								case INLINESTR:
									XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
									thisStr = rtsi.toString();
									break;
		
								case SSTINDEX:
									String sstIndex = value.toString();
									try {
										int idx = Integer.parseInt(sstIndex);
										XSSFRichTextString rtss = new XSSFRichTextString(sharedStringsTable.getEntryAt(idx));
										thisStr = rtss.toString();
									}
									catch (NumberFormatException ex) {
										logger.info("Failed to parse SST index '" + sstIndex + "': " + ex.toString());
									}
									break;
		
								case NUMBER:
									//thisStr= value.toString(); // unformatted numeric value
								  String n = value.toString();
									if (this.formatString != null && n.length() > 0)
										thisStr = formatter.formatRawCellContents(Double.parseDouble(n), this.formatIndex, this.formatString);
									else
										thisStr = n;
									break;
		
								default:
									thisStr = "(TODO: Unexpected type: " + nextDataType + ")";
									break;
								}
								
								// counting the number of header columns. Header columns cannot be empty hence empty checking logic is not required.
								if(isHeader){
									++numberOfHeaders;
								}
								
								
					           
					             for (int i = lastColumnNumber; i < thisColumn-1; ++i){
					                 updateCurrentRow("");
					             }
				                 updateCurrentRow(thisStr);
		
					             // Update column
					             if (thisColumn > -1){
					                 lastColumnNumber = thisColumn;
					             }
							}else if ("row".equals(name)) {
					             // Print out any missing commas if needed for rows other than header row.
					             if (!isHeader && numberOfHeaders > 0) {
					                
					                 for (int i = lastColumnNumber; i < numberOfHeaders-1; i++) {
					                	 updateCurrentRow("");
					                 }
					             }
								
					            //ending header row
								if(isHeader){
									isHeader=false;
								}
								currentRowEnds=true;
								lastColumnNumber = -1;
								return this.currentRow;
							}
					    	
					    }
						
						if(streamReader.hasNext()){
							streamReader.next();
						}else{
							currentRowEnds=true;
						}
					}
			} catch (XMLStreamException e) {
				e.printStackTrace();
				throw new RuntimeException(e);
			}
			
			return null;
		}
		
		/**
		 * NOT SUPPORTED IN THE CURRENT IMPLEMENTATION.
		 * @author Dickens Prabhu
		 */
		public void remove() {
			throw new UnsupportedOperationException();
		}
		
		/**
		 * Converts an Excel column name like "C" to a zero-based index.
		 *
		 * @param name
		 * @return Index corresponding to the specified name
		 */
		private int nameToColumn(String name) {
			int column = -1;
			for (int i = 0; i < name.length(); ++i) {
				int c = name.charAt(i);
				column = (column + 1) * 26 + c - 'A';
			}
			return column;
		}
	}
	
}
