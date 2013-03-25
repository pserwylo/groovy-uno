package com.serwylo.uno.spreadsheet

import com.serwylo.uno.utils.OfficeFinder
import com.sun.star.sheet.XSpreadsheet
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.table.XCell
import com.sun.star.table.XCellRange

abstract class SpreadsheetTest extends GroovyTestCase {

	private SpreadsheetConnector connector

	protected void setUp() {
		super.setUp()
		// connector = new SpreadsheetConnector( OfficeFinder.createFinder( "/home/pete/apps/" ) )
		connector = new SpreadsheetConnector( OfficeFinder.createFinder( "/home/pete/apps/openoffice.org3/program/soffice" ) )
	}

	protected SpreadsheetConnector getConnector() {
		connector
	}

	protected void assertMatches( XCellRange range, String[][] values ) {
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
