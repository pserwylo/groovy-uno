package com.serwylo.uno.finder;

import java.io.File;

public class GenericOfficeFinder extends OfficeFinder {

	@Override
	protected String getExecutableName() {
		return "soffice";
	}

	@Override
	protected File getDefaultPath() {
		throw new RuntimeException( "Operating System '" + getOs() + "' not supported." );
	}

	@Override
	protected boolean isForRunningOs() {
		// This is the default fallback finder, so it should always at least get a run.
		return true;
	}

	@Override
	public File getInstalledDir() {
		throw new RuntimeException( "Operating System '" + getOs() + "' not supported." );
	}
}
