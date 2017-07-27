package com.dickens.core.parser;

import java.util.LinkedHashMap;
import java.util.Map;

import lombok.Data;


/**
 * Instantiates a new file record.
 * 
 * @author Dickens Prabhu
 */
@Data
public class FileRecord {
	
	/** The file record values. */
	private Map<String,String> fileRecordValues = new LinkedHashMap<String,String>();
	
	/** The row number. */
	private int rowNumber;
	
	
	/**
	 * Sets the value.
	 *
	 * @param columnName the column name
	 * @param value the value
	 */
	public void setValue(String columnName, String value) {
		fileRecordValues.put(columnName, value);
	}
	
	/**
	 * Gets the value.
	 *
	 * @param columnName the column name
	 * @return the value
	 */
	public String getValue(String columnName) {
		return fileRecordValues.get(columnName);
	}

}
