package com.serwylo.uno.spreadsheet

import com.serwylo.uno.DocumentOptions

abstract class SpreadsheetTestUsingSingleFile extends SpreadsheetTestUsingFile {


	protected String getPathTo( String filename ) {
		String resourceFilename = "/home/pete/code/groovy-uno/src/test/resources/docs/$testDocFolderName/$filename"
		return new File( resourceFilename )
		// TODO: SpreadsheetTest.class.getResource( resourceFilename ).file
	}

	protected void load( String filename, DocumentOptions options = null ) {
		super.load( getPathTo( filename ), options )
	}

	abstract protected String getTestDocFolderName()

}
