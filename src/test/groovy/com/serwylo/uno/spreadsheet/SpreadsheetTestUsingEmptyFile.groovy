package com.serwylo.uno.spreadsheet

import com.serwylo.uno.DocumentOptions
import com.sun.star.sheet.XSpreadsheetDocument

abstract class SpreadsheetTestUsingEmptyFile extends SpreadsheetTestUsingFile {

	protected void load() {
		super.load( null, null )
	}

}
