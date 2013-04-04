package com.serwylo.uno.finder;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

abstract public class OfficeFinder {

	public File searchDir;
	public int maxDepth = 3;

	public static OfficeFinder createFinder( File file ) {
		OfficeFinder finder = createFinder();
		finder.searchDir = file;
		return finder;
	}

	public static OfficeFinder createFinder( String path ) {
		return createFinder( new File( path ) );
	}

	public static OfficeFinder createFinder() {
		return createFinderForCurrentOs();
	}

	protected static OfficeFinder createFinderForCurrentOs() {
		OfficeFinder[] finders = {
			new LinuxOfficeFinder(),
			new WindowsOfficeFinder(),
			new MacOfficeFinder(),
			new GenericOfficeFinder()
		};
		OfficeFinder found = null;
		for ( OfficeFinder finder : finders ) {
			if ( finder.isForRunningOs() ) {
				found = finder;
				break;
			}
		}
		return found;
	}

	public File getSOfficePath() {
		File soffice;
		if ( searchDir != null ) {
			soffice = findInDir( searchDir, getExecutableName() );
		} else {
			soffice = getDefaultPath();
		}

		// ... and if the default one failed...
		if ( soffice == null || !soffice.exists() ) {
			throw new OfficeNotFoundException( searchDir, soffice, maxDepth );
		}

		return soffice;
	}

	/**
	 * If there was a findFileRecurse rather than just eachFileRecurse then it would be much easier, but as it is,
	 * I can't bust out of eachFileRecurse without an exception, so I'll just implement it here for now.
	 */
	protected File findInDir( File dir, String filename ) {
		return findInDir( dir, filename, 0 );
	}

	protected File findInDir( File dir, String filename, int depth ) {
		File found = null;

		List<File> childDirs = new ArrayList<File>();

		if ( !dir.exists() ) {
			return null;
		} else if ( dir.isFile() && dir.getName().equals( filename ) ) {
			found = dir;
		} else {

			for ( File child : dir.listFiles() ) {
				if ( child.isDirectory() ) {
					childDirs.add( child );
				} else {
					if ( child.getName().equals( filename ) ) {
						found = child;
						break;
					}
				}
			}

		}

		if ( found == null && ( depth < maxDepth || depth < 0 ) ) {
			for ( File childDir : childDirs ) {
				File f = findInDir( childDir, filename, depth + 1 );
				if ( f != null ) {
					found = f;
					break;
				}
			}
		}

		return found;
	}

	protected static String getOs() {
		return System.getProperty( "os.name" ).toLowerCase();
	}

	public File getJarJuh() {
		return findInDir( getInstalledDir(), "juh.jar" );
	}

	abstract protected String getExecutableName();

	abstract protected File getDefaultPath();

	abstract protected boolean isForRunningOs();

	public abstract File getInstalledDir();

}





