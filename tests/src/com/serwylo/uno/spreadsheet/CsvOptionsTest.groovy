package com.serwylo.uno.spreadsheet

import com.sun.star.sheet.XSpreadsheet
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.table.XCell
import com.sun.star.table.XCellRange

class CsvOptionsTest extends GroovyTestCase {

	private static final String[][] TEST_VALUES = [
		[ "First", "Second", "Third" ],
		[ "1", "1", "1" ],
		[ "2", "2", "2" ],
		[ "3", "3", "3" ],
		[ "4", "4", "4" ],
		[ "5", "5", "5" ]
	]

	protected SpreadsheetConnector connector

	protected void setUp() {
		super.setUp()
		connector = new SpreadsheetConnector()
	}

	void testFieldDelimiters() {
		fieldDelimiter( CsvOptions.COMMA, "comma-delimiters.csv", "44,34,76,1,,0,false,false" )
		fieldDelimiter( CsvOptions.TAB,   "tab-delimiters.csv",    "9,34,76,1,,0,false,false" )
		fieldDelimiter( CsvOptions.SPACE, "space-delimiters.csv", "32,34,76,1,,0,false,false" )
	}

	private XSpreadsheet load( String filename, CsvOptions options ) {
		println "$filename: $options"
		XSpreadsheetDocument doc = connector.open( "docs/csv/$filename", options )
		doc.sheets[ 0 ]
	}

	private void fieldDelimiter( char delimiter, String file, String filterOptions ) {
		CsvOptions options = new CsvOptions( fieldDelimiters: delimiter )
		assertEquals( "CSV Options", filterOptions, options.toString() )

		XSpreadsheet sheet = load( file, options )
		assertMatches( sheet["A1:C6"], TEST_VALUES )
	}

	private void assertMatches( XCellRange range, String[][] values ) {
		int colCount = range.address.EndColumn - range.address.StartColumn
		int rowCount = range.address.EndRow    - range.address.StartRow

		assertEquals( "Row count", rowCount, values.length - 1 )
		values.eachWithIndex { String[] row, int index ->
			assertEquals( "Col count at row $index", colCount, row.length - 1 )
		}

		for ( int col = range.address.StartColumn; col <= range.address.EndColumn; col ++ ) {
			for ( int row = range.address.StartRow; row <= range.address.EndRow; row ++ ) {
				XCell cell = range.getCellByPosition( col, row )
				String valueFromSheet = cell.formula
				String expectedValue  = values[ row - range.address.StartRow ][ col - range.address.StartColumn ]
				assertEquals( "Cell value", expectedValue, valueFromSheet )
			}
		}
	}
}
