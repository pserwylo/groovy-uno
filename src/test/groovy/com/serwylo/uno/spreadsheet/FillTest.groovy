package com.serwylo.uno.spreadsheet

import com.serwylo.uno.utils.ColumnUtils
import com.sun.star.table.CellAddress

class FillTest extends SpreadsheetTestUsingEmptyFile {

	void testFillDown() {
		load()
		sheet[ "A1" ] << 1
		sheet[ "A1:A9" ].fillDown()
		assertEquals( 9, sheet[ "A9" ].data.pop().pop() )
	}

}
