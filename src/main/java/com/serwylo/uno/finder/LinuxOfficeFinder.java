package com.serwylo.uno.finder;

import java.io.File;

public class LinuxOfficeFinder extends OfficeFinder {

	@Override
	protected String getExecutableName() {
		return "soffice";
	}

	@Override
	protected File getDefaultPath() {
		return new File( "/usr/bin/soffice" );
	}

	@Override
	protected boolean isForRunningOs() {
		return getOs().contains( "linux" );
	}

	@Override
	public File getInstalledDir() {
		return new File( "/usr/share/libreoffice" );
	}
}
