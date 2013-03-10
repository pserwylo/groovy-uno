package com.serwylo.uno.spreadsheet

import com.serwylo.uno.Connector
import com.sun.star.frame.FrameSearchFlag
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
		PropertyValue[] properties = null
		if ( options ) {
			properties = new PropertyValue[ 2 ]
			properties[ 0 ] = options.propertyFilterName
			properties[ 1 ] = options.propertyFilterOptions
		}
		openSpreadsheet( filename, properties )
	}

	public XSpreadsheetDocument open( String path ) {
		openSpreadsheet( path )
	}

	protected XSpreadsheetDocument openSpreadsheet( String path, PropertyValue[] properties = null ) {
		if ( properties == null ) {
			properties = new PropertyValue[0]
		}

		if ( path ) {
			path = new File( path ).absolutePath
			path = "file://$path"
		}

		XComponent component = loader.loadComponentFromURL( path, "_blank", FrameSearchFlag.ALL, properties )
		componentToSpreadsheetDocument( component )
	}

	private XSpreadsheetDocument componentToSpreadsheetDocument( XComponent component ) {
		UnoRuntime.queryInterface( XSpreadsheetDocument.class, component )
	}

}
