package com.dickens.core.parser;

/**
 Copyright 2005 Bytecode Pty Ltd.

 Licensed under the Apache License, Version 2.0 (the "License");
 you may not use this file except in compliance with the License.
 You may obtain a copy of the License at

 http://www.apache.org/licenses/LICENSE-2.0

 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
 */

import java.io.BufferedReader;
import java.io.Closeable;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

/**
 * A very simple CSV reader released under a commercial-friendly license.
 * 
 * @author 
 * 
 */
public class CSVReader extends GenericFileReader implements Closeable {

	private BufferedReader br;

	private boolean hasNext = true;

	private CSVParser parser;

	private int skipLines;

	private boolean linesSkiped;

	/**
	 * The default line to start reading.
	 */
	public static final int DEFAULT_SKIP_LINES = 0;

	/**
	 * Constructs CSVReader using a comma for the separator.
	 * 
	 * @param reader
	 *            the reader to an underlying CSV source.
	 */
	public CSVReader(Reader reader) {
		this(reader, CSVParser.DEFAULT_SEPARATOR, CSVParser.DEFAULT_QUOTE_CHARACTER, CSVParser.DEFAULT_ESCAPE_CHARACTER);
	}

	public CSVReader(String filePath) throws FileNotFoundException {
		this(new FileReader(filePath), CSVParser.DEFAULT_SEPARATOR, CSVParser.DEFAULT_QUOTE_CHARACTER, CSVParser.DEFAULT_ESCAPE_CHARACTER);
	}

	public CSVReader(String filePath, boolean readEmptyRow) throws FileNotFoundException {
		this(new FileReader(filePath), CSVParser.DEFAULT_SEPARATOR, CSVParser.DEFAULT_QUOTE_CHARACTER, CSVParser.DEFAULT_ESCAPE_CHARACTER);
	}


	public CSVReader(InputStream inputStream, boolean b) {
		this(new InputStreamReader(inputStream), CSVParser.DEFAULT_SEPARATOR, CSVParser.DEFAULT_QUOTE_CHARACTER, CSVParser.DEFAULT_ESCAPE_CHARACTER);
	}

	/**
	 * Constructs CSVReader with supplied separator.
	 * 
	 * @param reader
	 *            the reader to an underlying CSV source.
	 * @param separator
	 *            the delimiter to use for separating entries.
	 */
	public CSVReader(Reader reader, char separator) {
		this(reader, separator, CSVParser.DEFAULT_QUOTE_CHARACTER, CSVParser.DEFAULT_ESCAPE_CHARACTER);
	}

	/**
	 * Constructs CSVReader with supplied separator and quote char.
	 * 
	 * @param reader
	 *            the reader to an underlying CSV source.
	 * @param separator
	 *            the delimiter to use for separating entries
	 * @param quotechar
	 *            the character to use for quoted elements
	 */
	public CSVReader(Reader reader, char separator, char quotechar) {
		this(reader, separator, quotechar, CSVParser.DEFAULT_ESCAPE_CHARACTER, DEFAULT_SKIP_LINES, CSVParser.DEFAULT_STRICT_QUOTES);
	}

	/**
	 * Constructs CSVReader with supplied separator, quote char and quote handling
	 * behavior.
	 *
	 * @param reader
	 *            the reader to an underlying CSV source.
	 * @param separator
	 *            the delimiter to use for separating entries
	 * @param quotechar
	 *            the character to use for quoted elements
	 * @param strictQuotes
	 *            sets if characters outside the quotes are ignored
	 */
	public CSVReader(Reader reader, char separator, char quotechar, boolean strictQuotes) {
		this(reader, separator, quotechar, CSVParser.DEFAULT_ESCAPE_CHARACTER, DEFAULT_SKIP_LINES, strictQuotes);
	}

	/**
	 * Constructs CSVReader with supplied separator and quote char.
	 *
	 * @param reader
	 *            the reader to an underlying CSV source.
	 * @param separator
	 *            the delimiter to use for separating entries
	 * @param quotechar
	 *            the character to use for quoted elements
	 * @param escape
	 *            the character to use for escaping a separator or quote
	 */

	public CSVReader(Reader reader, char separator,
			char quotechar, char escape) {
		this(reader, separator, quotechar, escape, DEFAULT_SKIP_LINES, CSVParser.DEFAULT_STRICT_QUOTES);
	}

	/**
	 * Constructs CSVReader with supplied separator and quote char.
	 * 
	 * @param reader
	 *            the reader to an underlying CSV source.
	 * @param separator
	 *            the delimiter to use for separating entries
	 * @param quotechar
	 *            the character to use for quoted elements
	 * @param line
	 *            the line number to skip for start reading 
	 */
	public CSVReader(Reader reader, char separator, char quotechar, int line) {
		this(reader, separator, quotechar, CSVParser.DEFAULT_ESCAPE_CHARACTER, line, CSVParser.DEFAULT_STRICT_QUOTES);
	}

	/**
	 * Constructs CSVReader with supplied separator and quote char.
	 *
	 * @param reader
	 *            the reader to an underlying CSV source.
	 * @param separator
	 *            the delimiter to use for separating entries
	 * @param quotechar
	 *            the character to use for quoted elements
	 * @param escape
	 *            the character to use for escaping a separator or quote
	 * @param line
	 *            the line number to skip for start reading
	 */
	public CSVReader(Reader reader, char separator, char quotechar, char escape, int line) {
		this(reader, separator, quotechar, escape, line, CSVParser.DEFAULT_STRICT_QUOTES);
	}

	/**
	 * Constructs CSVReader with supplied separator and quote char.
	 * 
	 * @param reader
	 *            the reader to an underlying CSV source.
	 * @param separator
	 *            the delimiter to use for separating entries
	 * @param quotechar
	 *            the character to use for quoted elements
	 * @param escape
	 *            the character to use for escaping a separator or quote
	 * @param line
	 *            the line number to skip for start reading
	 * @param strictQuotes
	 *            sets if characters outside the quotes are ignored
	 */
	public CSVReader(Reader reader, char separator, char quotechar, char escape, int line, boolean strictQuotes) {
		this(reader, separator, quotechar, escape, line, strictQuotes, CSVParser.DEFAULT_IGNORE_LEADING_WHITESPACE);
	}

	/**
	 * Constructs CSVReader with supplied separator and quote char.
	 * 
	 * @param reader
	 *            the reader to an underlying CSV source.
	 * @param separator
	 *            the delimiter to use for separating entries
	 * @param quotechar
	 *            the character to use for quoted elements
	 * @param escape
	 *            the character to use for escaping a separator or quote
	 * @param line
	 *            the line number to skip for start reading
	 * @param strictQuotes
	 *            sets if characters outside the quotes are ignored
	 * @param ignoreLeadingWhiteSpace
	 *            it true, parser should ignore white space before a quote in a field
	 */
	public CSVReader(Reader reader, char separator, char quotechar, char escape, int line, boolean strictQuotes, boolean ignoreLeadingWhiteSpace) {
		this.br = new BufferedReader(reader);
		this.parser = new CSVParser(separator, quotechar, escape, strictQuotes, ignoreLeadingWhiteSpace);
		this.skipLines = line;
	}

	/**
	 * Reads the next line from the buffer and converts to a string array.
	 * 
	 * @return a string array with each comma-separated element as a separate
	 *         entry.
	 * 
	 * @throws IOException
	 *             if bad things happen during the read
	 */
	public String[] readNext(String nextLine) throws IOException {

		String[] result = null;
		do {
			if (!hasNext) {
				return result; // should throw if still pending?
			}
			String[] r = parser.parseLineMulti(nextLine);
			if (r.length > 0) {
				if (result == null) {
					result = r;
				} else {
					String[] t = new String[result.length+r.length];
					System.arraycopy(result, 0, t, 0, result.length);
					System.arraycopy(r, 0, t, result.length, r.length);
					result = t;
				}
			}
		} while (parser.isPending());
		return result;
	}

	/**
	 * Reads the next line from the file.
	 * 
	 * @return the next line from the file without trailing newline
	 * @throws IOException
	 *             if bad things happen during the read
	 */
	private String getNextLine() throws IOException {
		if (!this.linesSkiped) {
			for (int i = 0; i < skipLines; i++) {
				br.readLine();
			}
			this.linesSkiped = true;
		}
		String nextLine = br.readLine();
		if (nextLine == null) {
			hasNext = false;
		}
		return hasNext ? nextLine : null;
	}

	/* 
	 * Implement iterator of list<String>
	 */
	@Override
	public Iterator<List<String>> getIterator() {
		return new CSVIterator<List<String>>();
	}

	private class CSVIterator<T> implements Iterator<List<String>> {
		
		private String nextLine=null;
		
		public boolean hasNext() {
			try {
				nextLine = getNextLine();
			} catch (IOException e) {
				e.printStackTrace();
			}
			return nextLine==null ? false : true;
		}

		public List<String> next() {
			try {
				return Arrays.asList(readNext(nextLine));
			} catch (IOException e) {
				e.printStackTrace();
			}
			return null;
		}

		public void remove() {
			
		}


	}

	/**
	 * Closes the underlying reader.
	 * 
	 * @throws IOException if the close fails
	 */
	public void close() throws IOException{
		br.close();
	}

}

