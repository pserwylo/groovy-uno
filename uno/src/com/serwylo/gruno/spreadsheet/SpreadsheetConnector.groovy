package com.serwylo.gruno.spreadsheet

import com.serwylo.gruno.Connector
import com.sun.star.lang.XComponent
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.uno.UnoRuntime
import com.sun.star.beans.PropertyValue

class SpreadsheetConnector extends Connector {

	public SpreadsheetConnector() {
		super()
	}

	public XSpreadsheetDocument open() {
		openSpreadsheet( generateNewUrl( "scalc" ) )
	}

	public XSpreadsheetDocument open( String filename, CsvOptions options ) {
		openSpreadsheet( filename )
	}

	public XSpreadsheetDocument open( String path ) {
		openSpreadsheet( path )
	}

	protected XSpreadsheetDocument openSpreadsheet( String path, PropertyValue[] properties = null ) {
		if ( properties == null ) {
			properties = new PropertyValue[0]
		}
		XComponent component = loader.loadComponentFromURL( path, "_blank", 0, properties )
		componentToSpreadsheetDocument( component )
	}

	private XSpreadsheetDocument componentToSpreadsheetDocument( XComponent component ) {
		UnoRuntime.queryInterface( XSpreadsheetDocument.class, component )
	}

}
