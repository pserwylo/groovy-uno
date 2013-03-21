package com.serwylo.uno.spreadsheet

import com.sun.star.table.XCellRange

class FormulasTest extends SpreadsheetTestUsingEmptyFile {

	private static final List<List<Object>> TEST_DATA = [
		[ 1, 10 ],
		[ 2, 20 ],
		[ 3, 30 ],
		[ 4, 40 ],
	]

	void setUp() {
		super.setUp()
		load()
	}

	void testCalc() {
		int numSheets = doc.sheets.size()

		sheet[ "A1:B4" ] << TEST_DATA

		Calculator calc = new Calculator( doc, 0 )

		Object actualValue = calc.singleCalc( "=SUM(Sheet1.A1:A4)" )
		assertEquals( 10, actualValue )
		assertNumSheets( numSheets )

		actualValue = calc.singleCalc( "=SUM(Sheet1.A1:B4)" )
		assertEquals( 110, actualValue )
		assertNumSheets( numSheets )

	}

	private void assertNumSheets( int numSheets ) {
		assertEquals( "Temporary sheet(s) left over", numSheets, doc.sheets.size() )
	}

	void testMaxMany() {
		sheet[ "A1:B4" ] << TEST_DATA
		XCellRange range = sheet[ "A1:B4" ]
		Calculator calc = new Calculator( doc, 0 )
		List<Double> values = calc.maxForEachColumn( range )
		assertArrayEquals( [ 4, 40 ].toArray(), values.toArray() )
	}

	void testMax() {
		XCellRange range = sheet[ "A1:B4" ]
		range << TEST_DATA

		Calculator calc = new Calculator( doc, 0 )
		double actualMax = TEST_DATA.collect { it.max() }.max() as Double
		double max = calc.max( range )
		assertEquals( "Max calculation", actualMax, max )

		double sum = calc.sum( range )
		double actualSum = TEST_DATA.collect { it.sum() }.sum() as Double
		assertEquals( "Sum calculation", actualSum, sum )

		double average = calc.average( range )
		double actualAverage = actualSum / 8
		assertEquals( "Average calculation", actualAverage, average )

	}

	void testFormulaEvaluation() {
		sheet[ "A2:B5" ] << TEST_DATA
		XCellRange headerRange = sheet[ "A1:B1" ]
		List<String> expectedFormulas  = [ "=SUM(A2:A5)", "=SUM(B2:B5)" ]
		List<String> expectedSumValues = [ 10, 100 ]

		headerRange.formulas = [ expectedFormulas ]
		assertFormulasEvaluate( headerRange, expectedFormulas, expectedSumValues )

		// Test passing in an array of Strings instead of a list...
		String[][] expectedFormulasArrays = new String[ 1 ][]
		expectedFormulasArrays[ 0 ] = expectedFormulas.toArray()
		headerRange.formulasFromArrays = expectedFormulasArrays
	}

	protected void assertFormulasEvaluate( XCellRange range, def expectedFormulas, List expectedSumValues ) {
		List<String> actualFormulas  = range.formulas.pop()
		List<String> actualSumValues = range.data.pop()

		assertArrayEquals( expectedFormulas instanceof List ? expectedFormulas.toArray() : expectedFormulas,  actualFormulas.toArray()  )
		assertArrayEquals( expectedSumValues.toArray(), actualSumValues.toArray() )
	}

}
