package com.serwylo.uno.spreadsheet

import com.serwylo.uno.Connector
import com.serwylo.uno.DocumentOptions
import com.sun.star.frame.FrameSearchFlag
import com.sun.star.frame.XStorable
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

	public XSpreadsheetDocument open( String filename, DocumentOptions options = null ) {
		openSpreadsheet( filename, options?.properties )
	}

	protected XSpreadsheetDocument openSpreadsheet( String path, PropertyValue[] properties = null ) {
		if ( properties == null ) {
			properties = new PropertyValue[0]
		}
		path = sanitizePath( path )
		XComponent component = loader.loadComponentFromURL( path, "_blank", FrameSearchFlag.ALL, properties )
		componentToSpreadsheetDocument( component )
	}

	public void save( XSpreadsheetDocument document, String path, DocumentOptions options = null ) {
		saveSpreadsheet( document, path, options?.properties )
	}

	protected void saveSpreadsheet( XSpreadsheetDocument document, String path, PropertyValue[] properties = null ) {
		if ( properties == null ) {
			properties = new PropertyValue[0]
		}
		path = sanitizePath( path )
		XStorable store = UnoRuntime.queryInterface( XStorable.class, document )
		store.storeAsURL( path, properties )
	}

	private XSpreadsheetDocument componentToSpreadsheetDocument( XComponent component ) {
		UnoRuntime.queryInterface( XSpreadsheetDocument.class, component )
	}

}
