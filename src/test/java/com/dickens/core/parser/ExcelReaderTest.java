package com.dickens.core.parser;

import java.util.Iterator;
import java.util.List;

public class ExcelReaderTest {

	public static void main(String args[]){
		   GenericFileReader reader = null;
		try {
			reader = GenericFileReader.getReader("C:\\Users\\Dickens.Prabhu\\Downloads\\Database_fieldlist_V1.0.xlsx",true);
			   Iterator<List<String>> iterator = reader.getIterator(); 
			   while(iterator.hasNext()){                              
			 	List<String> row = iterator.next();  
	 			System.out.println(row);                     
			   }                                         
                               
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally{
			  try {
				  System.out.println("Done!!");
				reader.close();
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}  
	}
}
