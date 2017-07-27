
package com.dickens.core.parser;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.stream.XMLStreamException;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.springframework.web.multipart.MultipartFile;
import org.xml.sax.SAXException;

/**
 * The Class XLFileReader.
 * 
 * @author Dickens Prabhu
 */
public class XLFileReader {

	/**
	 * Reads the header of a excel file.
	 *
	 * @param filePath the file path
	 * @return the excel headers
	 */
	public static final Map<Integer, String> getExcelHeaders(String filePath) {
		Map<Integer, List<String>> rowValuesMap = XLFileReader.excelReader(filePath, 0, 1);

		List<String> columns = rowValuesMap.get(0);
		Map<Integer, String> columnMap = new LinkedHashMap<Integer, String>();
		int i = 0;
		for (String column : columns) {
			columnMap.put(++i, column);
		}

		return columnMap;
	}

	/**
	 * Reads the header of a excel file.
	 *
	 * @param file the file
	 * @return the excel headers
	 */
	public static final Map<Integer, String> getExcelHeaders(MultipartFile file) {
		Map<Integer, List<String>> rowValuesMap = XLFileReader.excelReader(file, 0, 1);

		List<String> columns = rowValuesMap.get(0);
		Map<Integer, String> columnMap = new LinkedHashMap<Integer, String>();
		int i = 0;
		for (String column : columns) {
			columnMap.put(i++, column);
		}

		return columnMap;
	}

	/**
	 * Reads all records of excel file.
	 *
	 * @param filePath            absolute path of file
	 * @return the map
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public static Map<Integer, List<String>> excelReader(String filePath) throws IOException {
		return excelReader(filePath, -1, -1);
	}

	/**
	 * Reads all records of excel file.
	 *
	 * @param file the file
	 * @return the map
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public static Map<Integer, List<String>> excelReader(MultipartFile file) throws IOException {
		return excelReader(file, -1, -1);
	}

	/**
	 * Read records with a given offset and limit.
	 *
	 * @param file the file
	 * @param offset            start row number
	 * @param limit            number of records to fetch
	 * @return the map
	 */
	public static Map<Integer, List<String>> excelReader(MultipartFile file, int offset, int limit) {
		GenericFileReader excelReader=null;
		try {
			excelReader = getExcelReader(file);
			return getRowValuesMap(offset, limit, excelReader);
		} catch (Exception e) {
			throw new RuntimeException("Exception in reading file", e);
		} finally {
			if(excelReader!=null) {
				try {
					excelReader.close();
				} catch (Exception e) {
					//log.error("Error in reading file.");
				}
			}
		}

	}

	/**
	 * Read records with a given offset and limit.
	 *
	 * @param filePath            absolute path of file
	 * @param offset            start row number
	 * @param limit            number of records to fetch
	 * @return the map
	 */
	public static Map<Integer, List<String>> excelReader(String filePath, int offset, int limit) {
		GenericFileReader excelReader=null;
		try {
			excelReader = getExcelReader(filePath);
			return getRowValuesMap(offset, limit, excelReader);
		} catch (Exception e) {
			throw new RuntimeException("Exception in reading file", e);
		} finally {
			if(excelReader!=null) {
				try {
					excelReader.close();
				} catch (Exception e) {
					//log.error("Error in reading file.", e);
				}
			}
		}

	}

	/**
	 * Process file.
	 *
	 * @param filePath the file path
	 * @param dataMapping the data mapping
	 * @param offset the offset
	 * @param limit the limit
	 * @return the list
	 */
	public static List<FileRecord> processFile(String filePath, List<ColumnsMap> dataMapping, int offset, int limit) {

		List<FileRecord> fileRecords = null;
		try {
			Map<Integer, List<String>> rowValuesMap = XLFileReader.excelReader(filePath, offset, limit);
			fileRecords = getFileRecords(rowValuesMap, dataMapping);
		} catch (Exception e) {
			//log.error("Error in reading file.", e);
			throw new RuntimeException(e);
		}
		//log.info("File (" + filePath + " ) has " + fileRecords.size() + " records.");
		return fileRecords;

	}


	
	/**
	 * Gets the row values map.
	 *
	 * @param offset the offset
	 * @param limit the limit
	 * @param excelReader the excel reader
	 * @return the row values map
	 */
	private static Map<Integer, List<String>> getRowValuesMap(int offset, int limit, GenericFileReader excelReader) {
		Map<Integer, List<String>> rowValuesMap = new LinkedHashMap<Integer, List<String>>();

		int rowNumber = 0;
		int count = 0;
		if (excelReader != null) {
			Iterator<List<String>> iterator = excelReader.getIterator();
			while (iterator.hasNext()) {
				List<String> values = iterator.next();
				if (offset == -1 || (rowNumber >= offset && (limit == -1 || count < limit))) {
					rowValuesMap.put(rowNumber, new ArrayList<String>(values));
					count++;
				}
				rowNumber++;

				if (limit != -1 && rowNumber > (offset + limit)) {
					break;
				}
			}
		}
		return rowValuesMap;
	}

	/**
	 * Gets the excel reader.
	 *
	 * @param filePath the file path
	 * @return the excel reader
	 * @throws IOException Signals that an I/O exception has occurred.
	 * @throws OpenXML4JException the open XML 4 J exception
	 * @throws ParserConfigurationException the parser configuration exception
	 * @throws SAXException the SAX exception
	 * @throws XMLStreamException the XML stream exception
	 */
	private static GenericFileReader getExcelReader(String filePath)
			throws IOException, OpenXML4JException, ParserConfigurationException, SAXException, XMLStreamException {
		GenericFileReader excelReader;
		if (filePath.endsWith(".xlsx")) {
			excelReader = new XLSXReader(filePath,false);
		} else if (filePath.endsWith(".csv")) {
			excelReader = new CSVReader(filePath,false);
		} else {
			excelReader = new XLSReader(filePath,false);
		}
		return excelReader;
	}

	/**
	 * Gets the excel reader.
	 *
	 * @param file the file
	 * @return the excel reader
	 * @throws IOException Signals that an I/O exception has occurred.
	 * @throws OpenXML4JException the open XML 4 J exception
	 * @throws ParserConfigurationException the parser configuration exception
	 * @throws SAXException the SAX exception
	 * @throws XMLStreamException the XML stream exception
	 */
	private static GenericFileReader getExcelReader(MultipartFile file)
			throws IOException, OpenXML4JException, ParserConfigurationException, SAXException, XMLStreamException {
		GenericFileReader excelReader;
		if (file.getOriginalFilename().endsWith(".xlsx")) {
			excelReader = new XLSXReader(file.getInputStream(),false);
		}else if (file.getOriginalFilename().endsWith(".csv")) {
			excelReader = new CSVReader(file.getInputStream(),false);
		} else {
			excelReader = new XLSReader(file.getInputStream(),false);
		}
		return excelReader;
	}

	/**
	 * Gets the file records.
	 *
	 * @param rowValuesMap the row values map
	 * @param columnMap the column map
	 * @return the file records
	 */
	private static List<FileRecord> getFileRecords(Map<Integer, List<String>> rowValuesMap,
			List<ColumnsMap> columnMap) {
		List<FileRecord> fileRecords = new ArrayList<FileRecord>();
		for (Map.Entry<Integer, List<String>> entryMap : rowValuesMap.entrySet()) {
			FileRecord fileRecord = new FileRecord();
			int columnIndex = 1;
			List<String> values = entryMap.getValue();

			for (String value : values) {
				for (ColumnsMap columns : columnMap) {
					if (columnIndex == columns.getColumnIndex()) { // means it
																	// is
																	// matching
																	// with the
																	// column
						fileRecord.setValue(columns.getMappedFieldName(), value);
					}
				}
				columnIndex++;
			}
			fileRecord.setRowNumber(entryMap.getKey());
			fileRecords.add(fileRecord);
		}

		return fileRecords;
	}

}
