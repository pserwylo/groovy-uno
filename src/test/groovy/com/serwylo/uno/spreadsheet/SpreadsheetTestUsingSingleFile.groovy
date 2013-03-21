package com.serwylo.uno.spreadsheet

import com.serwylo.uno.DocumentOptions

abstract class SpreadsheetTestUsingSingleFile extends SpreadsheetTestUsingFile {


	protected String getPathTo( String filename ) {
		String resourceFilename = "/docs/$testDocFolderName/$filename"
		this.class.getResource( resourceFilename ).file
	}

	protected void load( String filename, DocumentOptions options = null ) {
		super.load( getPathTo( filename ), options )
	}

	abstract protected String getTestDocFolderName()

}
