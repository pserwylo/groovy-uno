package com.serwylo.uno.utils

import groovy.io.FileType
import ooo.connector.server.OfficePath

abstract class OfficeFinder {

	public File searchDir

	public static OfficeFinder createFinder() {
		[
			new LinuxOfficeFinder(),
			new WindowsOfficeFinder(),
			new MacOfficeFinder(),
			new GenericOfficeFinder()
		].find { finder ->
			finder.isForRunningOs()
		}
	}

	public OfficePath getPath() {
		File soffice
		if ( searchDir != null ) {
			soffice = findInDir( searchDir )
		} else {
			soffice = getDefaultPath()
		}

		// ... and if the default one failed...
		if ( !soffice || !soffice.exists() ) {
			throw new OfficeNotFoundException( searchDir, soffice )
		}

		new OfficePath( soffice )
	}

	/**
	 * If there was a findFileRecurse rather than just eachFileRecurse then it would be much easier, but as it is,
	 * I can't bust out of eachFileRecurse without an exception, so I'll just implement it here for now.
	 * @param dir
	 * @return
	 */
	protected File findInDir( File dir ) {
		File soffice = null

		if ( !dir.exists() ) {
			return null
		}

		def validate = { file ->
			if ( file.name == executableName ) {
				soffice = file
			}
		}

		validate( dir )

		dir.eachFile( FileType.FILES ) { file ->
			validate( file )
		}

		if ( !soffice ) {
			dir.eachFile( FileType.DIRECTORIES ) { childDir ->
				soffice = findInDir( childDir )
			}
		}

		soffice
	}

	protected static String getOs() {
		System.getProperty( "os.name" ).toLowerCase()
	}

	abstract protected String getExecutableName()

	abstract protected File getDefaultPath()

	abstract protected boolean isForRunningOs()

}

class OfficeNotFoundException extends FileNotFoundException {

	public OfficeNotFoundException( File searchDir, File defaultPath ) {
		super(
			"Couldn't find soffice executable at " +
			[ searchDir, defaultPath ].findAll { it != null }.join( " or " )
		)
	}

}

class LinuxOfficeFinder extends OfficeFinder {

	@Override
	protected String getExecutableName() {
		"soffice"
	}

	@Override
	protected File getDefaultPath() {
		new File( "/usr/bin/soffice" )
	}

	@Override
	protected boolean isForRunningOs() {
		os.contains( "linux" )
	}
}

class WindowsOfficeFinder extends OfficeFinder {

	@Override
	protected String getExecutableName() {
		"soffice.exe"
	}

	@Override
	protected File getDefaultPath() {

		File soffice = null

		List<File> programDirs = findBeginningWith( new File( "C:" + File.separator ), [ "Program Files", "Program Files (x86)" ] )
		for ( File programFilesDir in programDirs ) {

			List<File> officeDirs = findBeginningWith( programFilesDir, [ "LibreOffice", "OpenOffice" ] )
			for ( File officeDir in officeDirs ) {

				File sofficeFile = new File( officeDir.absolutePath + File.separator + "program" + File.separator + executableName )
				if ( sofficeFile.exists() ) {
					soffice = sofficeFile
					break;
				}
			}

			if ( soffice ) {
				break;
			}
		}

		soffice
	}

	protected List<File> findBeginningWith( File parent, List<String> beginsWith ) {
		List<File> found = []
		if ( parent.exists() ) {
			beginsWith.each { beginning ->
				parent.eachFileMatch( FileType.DIRECTORIES, ~"$beginning.*" ) {
					found << it
				}
			}
		}
		return found
	}

	@Override
	protected boolean isForRunningOs() {
		os.contains( "win" )
	}
}

class MacOfficeFinder extends LinuxOfficeFinder {

	@Override
	protected File getDefaultPath() {
		throw new Exception( "Mac not supported.")
	}

	@Override
	protected boolean isForRunningOs() {
		os.contains( "mac" )
	}
}

class GenericOfficeFinder extends OfficeFinder {

	@Override
	protected String getExecutableName() {
		return "soffice"
	}

	@Override
	protected File getDefaultPath() {
		throw new Exception( "Operating System '$os' not supported.")
	}

	@Override
	protected boolean isForRunningOs() {
		// This is the default fallback finder, so it should always at least get a run.
		true
	}
}