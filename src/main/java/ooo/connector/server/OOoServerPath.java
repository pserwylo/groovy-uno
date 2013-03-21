package ooo.connector.server;

public class OOoServerPath {

	private String sOfficePath = null;
	private String programPath = null;

	public OOoServerPath setExecPath( String path ) {
		setSOfficePath( path );
		setProgramPath( path );
		return this;
	}

	public OOoServerPath setSOfficePath( String path ) {
		this.sOfficePath = path;
		return this;
	}

	public OOoServerPath setProgramPath( String path ) {
		this.programPath = path;
		return this;
	}

	public String getSOfficePath() {
		return this.sOfficePath;
	}

	public String getProgramPath() {
		return this.programPath;
	}

}
