package com.serwylo.uno.finder;

import java.io.File;

public class MacOfficeFinder extends LinuxOfficeFinder {

	@Override
	protected File getDefaultPath() {
		throw new RuntimeException( "Mac not supported.");
	}

	@Override
	protected boolean isForRunningOs() {
		return getOs().contains( "mac" );
	}
}

