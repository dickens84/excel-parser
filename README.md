# Excel Parser
    It is a excel file reader implemented using Iterator design pattern.
    It internally uses default POI APIs for reading excel files.
# Recommended Use
    GenericFileReader reader = GenericFileReader.getReader(filePath);  
    Iterator<List<String>> iterator = reader.getIterator(); 
    while(iterator.hasNext()){                              
 	    List<String> row = iterator.next();                     
 	    if(row!=null){                                          
 		    for(String data:row){                              
 			      // process data	here                            
 	      }                                                  
      }
    }
    reader.close(); 
    
    It is an abstract class having two concrete implementations:
      1) XLSReader - for reading xls files
      2) XLSXReader - for reading xlsx files   
# Dependency 
    <dependency>
	<groupId>com.dickens.core</groupId>
        <artifactId>excel-parser</artifactId>
        <version>0.0.1-SNAPSHOT</version>
    </dependency>
