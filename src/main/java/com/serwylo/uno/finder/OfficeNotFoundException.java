package com.serwylo.uno.finder;

import java.io.File;

public class OfficeNotFoundException extends RuntimeException {

	public OfficeNotFoundException( File searchDir, File defaultPath, int maxDepth ) {
		super(
			"Couldn't find soffice executable at " + searchDir + " or " + defaultPath +
			( maxDepth >= 0 ? " (searched " + maxDepth + "folders deep)" : "" )
		);
	}

}
