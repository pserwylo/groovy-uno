package ooo.connector.server;

import java.io.File;

public class OfficePath
{

	private File binaryFile = null;

	public OfficePath( File binary ) {
		this.binaryFile = binary;
	}

	public OfficePath( String binaryPath ) {
		this.binaryFile = new File( binaryPath );
	}

	public File getBinaryFile() {
		return this.binaryFile;
	}

	public OfficePath setBinaryFile( File file ) {
		this.binaryFile = file;
		return this;
	}

}
