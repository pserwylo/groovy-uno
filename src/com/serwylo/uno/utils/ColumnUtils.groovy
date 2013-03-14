package com.serwylo.uno.utils

/**
 * http://www.blackbeltcoder.com/Articles/strings/converting-between-integers-and-spreadsheet-column-labels
 */
class ColumnUtils {

	private static final String ALPHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

	public static String indexToName( int index ) throws IllegalArgumentException {
		StringBuilder builder = new StringBuilder()
		boolean keepGoing = true
		while( keepGoing )
		{
			if ( builder.size() > 0 )
				index --
			builder.append( ALPHA[ index % ALPHA.size() ] )
			index /= ALPHA.size();
			keepGoing = ( index > 0 )
		}
		return builder.reverse().toString()
	}

	/*public static Integer columnNameToIndex( String name ) throws IllegalArgumentException {

		name = name.toUpperCase()


		boolean isValid = true
		Integer column  = 0

		for ( int i = name.length() - 1; i>= 0 ; i -- ) {
			char letter = name[ i ]
			int  index  = validChars.indexOf( letter )
			if ( index >= 0 ) {
				int value = index + 1
			} else {
				isValid = false
				break;
			}
		}

	}*/

}
