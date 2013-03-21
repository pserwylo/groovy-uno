package com.serwylo.uno.spreadsheet

import com.serwylo.uno.utils.ColumnUtils
import com.sun.star.sheet.XSpreadsheet
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.table.XCellRange

class Calculator {

	private XSpreadsheetDocument doc
	private XSpreadsheet         sheet

	private static final String MIN     = "MIN"
	private static final String MAX     = "MAX"
	private static final String AVERAGE = "AVERAGE"
	private static final String SUM     = "SUM"
	private static final String COUNT   = "COUNT"

	public Calculator( XSpreadsheetDocument doc, int sheetIndex ) {
		this.doc = doc
		sheet    = doc[ sheetIndex ]
	}

	public Calculator( XSpreadsheetDocument doc, String sheetName ) {
		this.doc = doc
		sheet    = doc[ sheetName ]
	}

	private void tempWorkingSheet( Closure c ) {
		String tempName = "Temp sheet " + Math.random()
		short tempIndex = (short)doc.sheets.size()
		doc.sheets.insertNewByName( tempName, tempIndex )

		c( tempName )

		doc.sheets.removeByName( tempName )
	}

	protected Object singleCalc( String formula ) {
		Object value = null
		tempWorkingSheet { String sheetName ->
			XCellRange range = doc.sheets[ sheetName ][ "A1" ]
			range.setFormula( formula )
			value = range.getData().pop().pop()
		}
		return value
	}

	protected List<Object> multiCalc( List<String> formulas ) {
		List<Object> values = null
		tempWorkingSheet { String sheetName ->
			String rangeEnd = ColumnUtils.indexToName( formulas.size() - 1 )
			XCellRange range = doc.sheets[ sheetName ][ "A1:${rangeEnd}1" ]
			range.setFormulas( [ formulas ] )
			values = range.getData().pop()
		}
		return values
	}

	public List<Double> maxForEachColumn( XCellRange range ) {
		functionForEachColumn( MAX, range )
	}

	public List<Double> minForEachColumn( XCellRange range ) {
		functionForEachColumn( MIN, range )
	}

	public List<Double> sumForEachColumn( XCellRange range ) {
		functionForEachColumn( SUM, range )
	}

	public List<Double> averageForEachColumn( XCellRange range ) {
		functionForEachColumn( AVERAGE, range )
	}

	public List<Double> functionForEachColumn( String function, XCellRange range ) {
		int startRow = range.address.StartRow + 1
		int endRow = range.address.EndRow + 1
		List<String> formulas = []
		for ( int i = range.address.StartColumn; i <= range.address.EndColumn; i ++ ) {
			String colName = ColumnUtils.indexToName( i )
			String address = "$sheet.name.$colName$startRow:$colName$endRow"
			formulas.add( "=$function( $address )" )
		}
		return multiCalc( formulas )
	}

	public Double max( XCellRange range ) {
		function( MAX, range )
	}

	public Double min( XCellRange range ) {
		function( MIN, range )
	}

	public Double average( XCellRange range ) {
		function( AVERAGE, range )
	}

	public Double sum( XCellRange range ) {
		function( SUM, range )
	}

	public Double function( String function, XCellRange range ) {
		String address = "$sheet.name.$range.name"
		singleCalc( "=$function( $address )" ) as Double
	}


}
