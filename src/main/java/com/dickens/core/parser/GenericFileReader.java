package com.dickens.core.parser;

import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.stream.XMLStreamException;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

/**
 * It is an abstract class having two concrete implementations:
 * 1) XLSReader - for reading xls files
 * 2) XLSXReader - for reading xlsx files
 * 
 * Depending upon the file extension, one of the above concrete
 * class is used to process the excel file.
 * 
 ***********************************************************
 ################## Recommended Use: ########################
 #  GenericFileReader reader = GenericFileReader.getReader(filePath);   #
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
public abstract class GenericFileReader {
	public abstract Iterator<List<String>> getIterator();
	public abstract void close() throws Exception;
	
	/**
	 * This method checks the extension of the file to be read.
	 * if file extension is .xls it returns XLSReader object.
	 * If file extension is .xlsx it return XLSXReader object.
	 * 
	 * @author Dickens Prabhu
	 * @param filePath
	 * @return  
	 * @throws IOException
	 * @throws OpenXML4JException
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 * @throws XMLStreamException
	 */
	public static GenericFileReader getReader(String filePath) throws IOException, OpenXML4JException, ParserConfigurationException, SAXException, XMLStreamException{
		String extension = FilenameUtils.getExtension(filePath);
		if(extension==null){
			return null;
		}
		if("xlsx".equalsIgnoreCase(extension)){
			return new XLSXReader(filePath);
		}else if("csv".equalsIgnoreCase(extension)){
			return new CSVReader(filePath);
		}else if("xls".equalsIgnoreCase(extension)){
			return new XLSReader(filePath);
		}
		
		return null;
	}
	
	/**
	 * This method checks the extension of the file to be read.
	 * if file extension is .xls it returns XLSReader object.
	 * If file extension is .xlsx it return XLSXReader object.
	 * 
	 * @author Dickens Prabhu
	 * @param filePath
	 * @return  
	 * @throws IOException
	 * @throws OpenXML4JException
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 * @throws XMLStreamException
	 */
	public static GenericFileReader getReader(String filePath, boolean readEmptyRow) throws IOException, OpenXML4JException, ParserConfigurationException, SAXException, XMLStreamException{
		String extension = FilenameUtils.getExtension(filePath);
		if(extension==null){
			return null;
		}
		
		if("xlsx".equalsIgnoreCase(extension)){
			return new XLSXReader(filePath,readEmptyRow);
		}else if("csv".equalsIgnoreCase(extension)){
			return new CSVReader(filePath, readEmptyRow);
		}else if("xls".equalsIgnoreCase(extension)){
			return new XLSReader(filePath,readEmptyRow);
		}
		
		return null;
	}
	
	/**
	 * Returns false if even one of the string in the passed list contains some data(other than empty string or white spaces).
	 * Otherwise it returns true.
	 * 
	 * @author Dickens Prabhu
	 * @param currentRow
	 * @return
	 */
	protected boolean isEmptyCurrentRow(List<String> currentRow) {
		if(currentRow==null || currentRow.isEmpty()){
			return true;
		}
		for(String data:currentRow){
			if(!(StringUtils.isEmpty(data) || StringUtils.isWhitespace(data))){ // check even one of the data in the current row is not [empty,whitespace]
				return false; // current row is not empty
			}
		}
		return true; // current row is empty
	}
}
