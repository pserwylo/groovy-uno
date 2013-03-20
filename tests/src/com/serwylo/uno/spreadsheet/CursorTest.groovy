package com.serwylo.uno.spreadsheet

import com.serwylo.uno.utils.ColumnUtils
import com.sun.star.sheet.XSpreadsheet
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.table.CellAddress

class CursorTest extends SpreadsheetTest {

	void testMaxRange() {
		assertDocumentDimensions( "docs/cursor-Z1000.ods", "Z", 1000 )
		assertDocumentDimensions( "docs/cursor-Z500-H1000.ods", "Z", 1000 )
	}

	private void assertDocumentDimensions( String document, String maxColumn, int maxRow ) {
		XSpreadsheetDocument doc = connector.open( document )
		XSpreadsheet sheet       = doc[ 0 ]

		CellAddress dimensions = sheet.dimensions
		assertEquals( maxColumn, ColumnUtils.indexToName( dimensions.Column ) )
		assertEquals( maxRow, dimensions.Row + 1 )

		doc.close()
	}

}
