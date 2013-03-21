package com.serwylo.uno.spreadsheet

import com.serwylo.uno.DocumentOptions
import com.sun.star.sheet.XSpreadsheet
import com.sun.star.sheet.XSpreadsheetDocument

abstract class SpreadsheetTestUsingFile extends SpreadsheetTest {

	private XSpreadsheetDocument doc
	private XSpreadsheet sheet

	protected void load( String filename, DocumentOptions options = null ) {
		unload()
		if ( filename ) {
			doc = connector.open( filename, options )
		} else {
			doc = connector.open()
		}
		sheet = doc[ 0 ]
	}

	protected void unload() {
		doc?.close()
		doc   = null
		sheet = null
	}

	protected void tearDown() {
		unload()
		super.tearDown()
	}

	protected XSpreadsheetDocument getDoc() {
		doc
	}

	protected XSpreadsheet getSheet() {
		sheet
	}

}
