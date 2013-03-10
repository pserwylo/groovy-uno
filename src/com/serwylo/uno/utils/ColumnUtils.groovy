package com.serwylo.uno.utils


class ColumnUtils {

	public static Integer columnNameToIndex( String name ) throws IllegalArgumentException {

		name = name.toUpperCase()
		String validChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

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

	}

}
