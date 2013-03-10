package com.serwylo.uno.spreadsheet

import com.sun.star.beans.PropertyValue

/**
 * @see http://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Filter_Options
 * http://forum.openoffice.org/en/forum/viewtopic.php?f=44&t=14018&p=65702&hilit=showfilteroptions
 * @see com/sun/star/document/MediaDescriptor.html
 */
class CsvOptions {

	private filterName = "Text - txt - csv (StarCalc)"

	/**
	 *            32,0,60,1,,0,false,true   // "Detect special numbers"
	 *            32,0,60,1,,0,true,false   // "Quoted text as field"
	 *        32/MRG,0,60,1,,0,false,false  // "Merge delimiters"
	 *            32,0,60,1,,0,false,false  // text()
	 *            32,39,60,1,,0,false,false // text(')
	 *            32,34,60,2,,0,false,false // Space delimiter, text("), From line 2
	 *            32,34,60,1,,0,false,false // Space delimiter, From line 1
	 * 9/59/44/32/48,34,60,1,,0,false,false // Tab/Comma/Other(0),Semicolon,Space delimiter
	 *            32,0,76,1,,0,false,false  // UTF-8 (others were some japanese encoding)
	 *            32,0,29,1,,0,false,false  // Arabic (DOS/OS2-864) - the first encoding in the list
	 *
	 *
	 */

	static final char TAB       = "\t"
	static final char COMMA     = ","
	static final char SEMICOLON = ";"
	static final char SPACE     = " "

	static final char DOUBLE_QUOTE = '"'
	static final char SINGLE_QUOTE = "'"

	boolean detectSpecialNumbers = false
	boolean quotedTextAsField    = false
	boolean mergeDelimiters      = false

	Character textDelimiter      = DOUBLE_QUOTE
	String fieldDelimiters       = TAB

	int startLine                = 0

	PropertyValue getPropertyFilterName() {
		new PropertyValue( Name: "FilterName", Value: "Text - txt - csv (StarCalc)" )
	}

	PropertyValue getPropertyFilterOptions() {
		new PropertyValue( Name: "FilterOptions", Value: toString() )
	}

	String toString() {
		[
			paramFieldDelimiters(),
			paramTextDelimiter(),
			paramEncoding(),
			paramStartLine(),
			paramMysticalProperty1(),
			paramMysticalProperty2(),
			paramQuotedTextAsField(),
			paramDetectSpecialNumbers(),
		].join( "," )
	}

	protected String paramFieldDelimiters() {
		String parts = fieldDelimiters.collect { (int)it }.join( "/" )
		if ( mergeDelimiters ) {
			parts += "/MRG"
		}
		return parts
	}

	protected String paramTextDelimiter() {
		(int)textDelimiter
	}

	protected String paramMysticalProperty1() {
		""
	}

	protected String paramMysticalProperty2() {
		0
	}

	/**
	 * At this stage, I cbf working out the different encoding options provided in their cryptic ways.
	 * So 76 just means UTF-8.
	 * @return
	 */
	protected String paramEncoding() {
		76
	}

	/**
	 * We use zero based counting, but UNO uses 1 based counting.
	 * @return
	 */
	protected String paramStartLine() {
		startLine + 1
	}

	protected String paramQuotedTextAsField() {
		quotedTextAsField ? "true" : "false"
	}

	protected String paramDetectSpecialNumbers() {
		detectSpecialNumbers ? "true" : "false"
	}

}
