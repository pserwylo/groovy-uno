package com.serwylo.uno.utils

import com.serwylo.uno.spreadsheet.SpreadsheetTest
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.table.XCellRange

class UtilsTest extends SpreadsheetTest {

	public void testIndexToColumnName() {

		def values = [
			"A"   : 0,   "B"   : 1,   "C"   : 2,
			"AA"  : 26,  "AB"  : 27,  "AC"  : 28,
			"BA"  : 52,  "BB"  : 53,  "BC"  : 54,
			"ZA"  : 676, "ZB"  : 677, "ZC"  : 678,
			"AAA" : 702, "AAB" : 703, "AAC" : 704,
		]

		values.each {
			assertEquals( it.key, ColumnUtils.indexToColumnName( it.value ) )
		}

	}

	public void testValueToArrays() {
		Double[][] arrayValues = [
			[ Math.PI, Math.PI, Math.PI ],
			[ Math.PI, Math.PI, Math.PI ]
		]

		Object[][] values = DataUtils.valueToArrays( Math.PI, 3, 2 )
		assertMatches( arrayValues, values )

		XSpreadsheetDocument doc = connector.open()
		XCellRange range = doc[ 0 ][ "A1:C2" ]
		values = DataUtils.valueToArrays( Math.PI, range )
		assertMatches( arrayValues, values )
		doc.close()
	}

	public void testListsToArrays() {
		String[][] arrayValues = [
			[ "First", "Second", "Third" ],
			[ "1", "1", "1" ],
			[ "2", "2", "2" ],
			[ "3", "3", "3" ],
			[ "4", "4", "4" ],
			[ "5", "5", "5" ]
		]

		List<List<Object>> listValues = [
			[ "First", "Second", "Third" ],
			[ "1", "1", "1" ],
			[ "2", "2", "2" ],
			[ "3", "3", "3" ],
			[ "4", "4", "4" ],
			[ "5", "5", "5" ]
		]

		Object[][] arrays = DataUtils.listsToArrays( listValues )
		assertMatches( arrays, arrayValues )
	}

	protected void assertMatches( Object[][] expected, Object[][] actual ) {
		assertEquals( "Num rows", expected.length, actual.length )
		for ( int i in 0..(expected.length-1) ) {
			Object[] expectedRow = expected[ i ]
			Object[] actualRow   = actual[ i ]
			assertEquals( "Num cols for row $i", expectedRow.length, actualRow.length )
			assertArrayEquals( expectedRow, actualRow )
		}
	}

}
