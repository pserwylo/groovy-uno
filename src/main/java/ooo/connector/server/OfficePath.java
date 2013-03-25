package ooo.connector.server;

import java.io.File;

public class OfficePath
{

	private File binaryFile = null;

	public OfficePath( File file ) {
		this.binaryFile = file;
	}

	public OfficePath( String path ) {
		this.binaryFile = new File( path );
	}

	public File getBinaryFile() {
		return this.binaryFile;
	}

	public OfficePath setBinaryFile( File file ) {
		this.binaryFile = file;
		return this;
	}

}
