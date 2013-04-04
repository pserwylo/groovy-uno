package com.serwylo.uno.finder;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class WindowsOfficeFinder extends OfficeFinder {

	@Override
	protected String getExecutableName() {
		return "soffice.exe";
	}

	private List<File> getPotentialOfficeDirs() {
		List<File> officeDirs = new ArrayList<File>();

		String[] potentialProgramFiles = { "Program Files", "Program Files (x86)" };
		String[] potentialOfficeFiles  = { "LibreOffice", "OpenOffice" };

		List<File> programFiles = findBeginningWith( new File( "C:" + File.separator ), potentialProgramFiles );
		for ( File dir : programFiles ) {
			officeDirs.addAll(findBeginningWith(dir, potentialOfficeFiles));
		}

		return officeDirs;
	}

	@Override
	protected File getDefaultPath() {

		File soffice = null;

		for ( File officeDir : getPotentialOfficeDirs() ) {
			File sofficeFile = new File( officeDir.getAbsolutePath() + File.separator + "program" + File.separator + getExecutableName() );
			if ( sofficeFile.exists() ) {
				soffice = sofficeFile;
				break;
			}
		}

		return soffice;
	}

	protected List<File> findBeginningWith( File parent, String[] beginsWith ) {
		List<File> found = new ArrayList<File>();
		if ( parent.exists() ) {
			for ( String start : beginsWith ) {
				for ( File child : parent.listFiles() ) {
					if ( child.getName().startsWith( start ) ) {
						found.add( child );
					}
				}
			}
		}
		return found;
	}

	@Override
	protected boolean isForRunningOs() {
		return getOs().contains( "win" );
	}

	@Override
	public File getInstalledDir() {
		// Parent file is the "program" directory, and its parent is the install dir...
		return getSOfficePath().getParentFile().getParentFile();
	}
}
