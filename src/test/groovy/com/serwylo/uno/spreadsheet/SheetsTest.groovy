package com.serwylo.uno.spreadsheet

import com.sun.star.sheet.XSpreadsheet
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.sheet.XSpreadsheets
import org.apache.ivy.util.ChecksumHelper

class SheetsTest extends SpreadsheetTestUsingSingleFile {

	private static final List<String> SHEET_NAMES        = [ "First sheet", "Second sheet", "Third sheet" ]
	private static final String MULTIPLE_SHEETS_FILENAME = "multiple-sheets.ods"

	void testGetAt() {
		load( MULTIPLE_SHEETS_FILENAME )
		SHEET_NAMES.each { name ->

			XSpreadsheet sheet1 = doc.sheets[ name ]
			XSpreadsheet sheet2 = doc[ name ]
			String value1 = sheet1.getCellByPosition( 0, 0 ).formula
			String value2 = sheet2.getCellByPosition( 0, 0 ).formula

			assertEquals( "Sheet obtained from doc.sheets[$name]", name, value1 )
			assertEquals( "Sheet obtained from doc[$name]", name, value2 )
			assertEquals( "Sheet name", name, sheet1.name )
			assertEquals( "Sheet name", name, sheet2.name )
		}
	}

	@Override
	protected String getTestDocFolderName() {
		"sheets"
	}
}
