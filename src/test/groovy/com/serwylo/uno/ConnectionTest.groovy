package com.serwylo.uno

import com.serwylo.uno.utils.LinuxOfficeFinder
import com.serwylo.uno.utils.MacOfficeFinder
import com.serwylo.uno.utils.OfficeFinder
import com.serwylo.uno.utils.WindowsOfficeFinder
import ooo.connector.server.OfficePath

class ConnectionTest extends GroovyTestCase {

	private static final String OS_WINDOWS = "Windows"
	private static final String OS_LINUX   = "Linux"
	private static final String OS_MAC     = "Mac"

	private String originalOsName

	private File tempDir

	public void setUp() {
		originalOsName = System.getProperty( "os.name" )
		tempDir = File.createTempDir( "groovy-uno-connection-test", "" )
	}

	public void tearDown() {
		tempDir.deleteDir()
		setOs( originalOsName )
	}

	public void testLinuxFinder() {
		setOs( OS_LINUX )
		OfficeFinder finder = OfficeFinder.createFinder()
		assertTrue( finder instanceof LinuxOfficeFinder )

		File fakeDir     = fakeLibreOfficeDir( "some" + File.separator + "folder" + File.separator + "LibreOffice 4.0" )
		finder.searchDir = fakeDir
		OfficePath path = finder.path
		assertEquals( "Office path", path.binaryFile.name, "soffice" )

	}

	/**
	 * TODO: Find way to actually test this in some way.
	 * The problem is that if I want to have a pseudo-chroot environment to play with.
	 */
	public void testWindowsFinder() {
		setOs( OS_WINDOWS )
		OfficeFinder finder = OfficeFinder.createFinder()
		assertTrue( finder instanceof WindowsOfficeFinder )
	}

	/**
	 * TODO: Find way to actually test this in some way.
	 * The problem is that if I want to have a pseudo-chroot environment to play with.
	 */
	public void testMacFinder() {
		setOs( OS_MAC )
		OfficeFinder finder = OfficeFinder.createFinder()
		assertTrue( finder instanceof MacOfficeFinder )
	}

	protected void setOs( String os ) {
		System.setProperty( "os.name", os )
		assertEquals( "Changed OS name.", os, System.getProperty( "os.name" ) )
	}

	protected File fakeLibreOfficeDir( String path ) {
		File dir = new File( tempDir, path )
		dir.mkdir()

		File programDir = new File( dir, "program" )
		programDir.mkdir()

		File soffice = new File( programDir, "soffice" )
		soffice.createNewFile()

		dir
	}

}
