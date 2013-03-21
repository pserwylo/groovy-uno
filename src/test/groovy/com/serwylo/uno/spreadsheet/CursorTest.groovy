package com.serwylo.uno.spreadsheet

import com.serwylo.uno.utils.ColumnUtils
import com.sun.star.table.CellAddress

class CursorTest extends SpreadsheetTestUsingSingleFile {

	void testMaxRange() {
		assertDocumentDimensions( "cursor-Z1000.ods", "Z", 1000 )
		assertDocumentDimensions( "cursor-Z500-H1000.ods", "Z", 1000 )
	}

	private void assertDocumentDimensions( String document, String maxColumn, int maxRow ) {
		load( document )
		CellAddress dimensions = sheet.dimensions
		assertEquals( maxColumn, ColumnUtils.indexToName( dimensions.Column ) )
		assertEquals( maxRow, dimensions.Row + 1 )
	}

	@Override
	protected String getTestDocFolderName() {
		"cursor"
	}
}
