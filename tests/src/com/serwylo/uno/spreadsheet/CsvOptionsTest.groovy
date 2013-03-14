package com.serwylo.uno.spreadsheet

import com.sun.star.sheet.XSpreadsheet
import com.sun.star.sheet.XSpreadsheetDocument
import org.apache.ivy.util.ChecksumHelper

class CsvOptionsTest extends SpreadsheetTest {

	private static final String[][] FIELD_DELIMITERS_VALUES = [
		[ "First", "Second", "Third" ],
		[ "1", "1", "1" ],
		[ "2", "2", "2" ],
		[ "3", "3", "3" ],
		[ "4", "4", "4" ],
		[ "5", "5", "5" ]
	]

	private static final String[][] TEXT_DELIMITER_VALUES = [
		[ "First", "Second", "Third" ],
		[ "1", 'Has "Double" quotes', "1" ],
		[ "2", "Has 'Single' quotes", "2" ],
		[ "3", "3", "3" ],
		[ "4", "4", "4" ],
		[ "5", "5", "5" ]
	]

	private static final CsvOptions OPT_TEXT_DOUBLE_QUOTE = new CsvOptions( textDelimiter   : CsvOptions.DOUBLE_QUOTE )
	private static final CsvOptions OPT_TEXT_SINGLE_QUOTE = new CsvOptions( textDelimiter   : CsvOptions.SINGLE_QUOTE )
	private static final CsvOptions OPT_FIELD_COMMA       = new CsvOptions( fieldDelimiters : CsvOptions.COMMA )
	private static final CsvOptions OPT_FIELD_TAB         = new CsvOptions( fieldDelimiters : CsvOptions.TAB )
	private static final CsvOptions OPT_FIELD_SPACE       = new CsvOptions( fieldDelimiters : CsvOptions.SPACE )

	private static final Map<CsvOptions, String> OPTIONS_TO_FILE = [
		(OPT_TEXT_DOUBLE_QUOTE) : "double-quotes-delimited.csv",
		(OPT_TEXT_SINGLE_QUOTE) : "single-quotes-delimited.csv",
		(OPT_FIELD_COMMA)       : "comma-delimiters.csv",
		(OPT_FIELD_TAB)         : "tab-delimiters.csv",
		(OPT_FIELD_SPACE)       : "space-delimiters.csv",
	]

	protected SpreadsheetConnector connector

	protected void setUp() {
		super.setUp()
		connector = new SpreadsheetConnector()
	}

	void testGetAt() {
		String file              = getPathFor( OPT_FIELD_COMMA )
		XSpreadsheetDocument doc = load( file, OPT_FIELD_COMMA )
		XSpreadsheet sheet       = doc[ 0 ]
		for ( int rowIndex = 0; rowIndex < FIELD_DELIMITERS_VALUES.length; rowIndex ++ ) {
			String[] row = FIELD_DELIMITERS_VALUES[ rowIndex ]
			for ( int columnIndex = 0; columnIndex < row.length; columnIndex ++ ) {
				String expected = row[ columnIndex ]
				String actual   = sheet[ columnIndex, rowIndex ].formula
				assertEquals( "Value at Column $columnIndex and Row $rowIndex", expected, actual )
			}
		}
		doc.close();
	}

	void testSave() {
		convertDocument( OPT_FIELD_COMMA, OPT_FIELD_TAB )
		convertDocument( OPT_FIELD_COMMA, OPT_FIELD_SPACE )
		convertDocument( OPT_FIELD_TAB, OPT_FIELD_COMMA )
		convertDocument( OPT_FIELD_TAB, OPT_FIELD_SPACE )
		convertDocument( OPT_FIELD_SPACE, OPT_FIELD_COMMA )
		convertDocument( OPT_FIELD_SPACE, OPT_FIELD_TAB )
		convertDocument( OPT_TEXT_DOUBLE_QUOTE, OPT_TEXT_SINGLE_QUOTE )
		convertDocument( OPT_TEXT_SINGLE_QUOTE, OPT_TEXT_DOUBLE_QUOTE )
	}

	protected void convertDocument( CsvOptions fromOptions, CsvOptions toOptions ) {
		String fromPath = getPathFor( fromOptions )
		String toPath   = getPathFor( toOptions   )

		// TODO: Temp file properly...
		File tempToFile = new File( "/tmp/${(int)(Math.random() * 10000)}-$toPath" )
		tempToFile.deleteOnExit()

		XSpreadsheetDocument doc = load( fromPath, fromOptions )
		connector.save( doc, tempToFile.path, toOptions )
		doc.close()

		String expected = ChecksumHelper.computeAsString( new File( getRelativePathToCsv( toPath ) ), "md5" )
		String actual   = ChecksumHelper.computeAsString( tempToFile, "md5" )

		assertEquals( "Converting from $fromPath to $toPath", expected, actual )
	}

	void testMergeDelimiters() {
		CsvOptions options = new CsvOptions( fieldDelimiters: CsvOptions.SPACE + CsvOptions.COMMA, mergeDelimiters: true )
		assertEquals( "CSV Options", "32/44/MRG,34,76,1,,0,false,false", options.toString() )

		XSpreadsheetDocument doc = load( "comma-and-space-delimiters.csv", options )
		assertMatches( doc[ 0 ]["A1:C6"], FIELD_DELIMITERS_VALUES )
		doc.close()
	}

	void testTextDelimiters() {
		textDelimiter( OPT_TEXT_DOUBLE_QUOTE, "44,34,76,1,,0,false,false" )
		textDelimiter( OPT_TEXT_SINGLE_QUOTE, "44,39,76,1,,0,false,false" )
	}

	void testFieldDelimiters() {
		fieldDelimiter( OPT_FIELD_COMMA, "44,34,76,1,,0,false,false" )
		fieldDelimiter( OPT_FIELD_TAB,   "9,34,76,1,,0,false,false" )
		fieldDelimiter( OPT_FIELD_SPACE, "32,34,76,1,,0,false,false" )
	}

	private XSpreadsheetDocument load( String filename, CsvOptions options ) {
		connector.open( getRelativePathToCsv( filename ), options )
	}

	private String getRelativePathToCsv( String filename ) {
		"docs/csv/$filename"
	}

	private String getPathFor( CsvOptions options ) {
		if ( !OPTIONS_TO_FILE.containsKey( options ) ) {
			throw new IllegalArgumentException( "Must be one of OPT_* constants." )
		}
		OPTIONS_TO_FILE.get( options )
	}

	private void textDelimiter( CsvOptions options, String filterOptions ) {
		assertEquals( "CSV Options", filterOptions, options.toString() )
		String file = getPathFor( options )
		XSpreadsheetDocument doc = load( file, options )
		assertMatches( doc[ 0 ]["A1:C6"], TEXT_DELIMITER_VALUES )
		doc.close()
	}

	private void fieldDelimiter( CsvOptions options, String filterOptions ) {
		assertEquals( "CSV Options", filterOptions, options.toString() )
		String file = getPathFor( options )
		XSpreadsheetDocument doc = load( file, options )
		assertMatches( doc[ 0 ]["A1:C6"], FIELD_DELIMITERS_VALUES )
		doc.close()
	}

}
