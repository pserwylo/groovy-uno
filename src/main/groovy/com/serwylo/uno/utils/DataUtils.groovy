package com.serwylo.uno.utils

import com.sun.star.table.CellRangeAddress
import com.sun.star.table.XCellRange


class DataUtils {


	public static String[][] stringListsToArrays( List<List<String>> lists ) {
		return genericListsToArrays( lists )
	}

	public static Double[][] doubleListsToArrays( List<List<Double>> lists ) {
		return genericListsToArrays( lists )
	}

	private static <T> T[][] genericListsToArrays( List<List<T>> lists ) {
		T[][] arrays = new T[ lists.size() ][]
		for ( int i in 0..( lists.size() - 1 ) ) {
			T[] row
			if ( lists[ i ] instanceof List ) {
				row = ( T[] ) lists[ i ].toArray()
			} else {
				row = new T[ 1 ]
				row[ 0 ] = lists[ i ]
			}
			arrays[ i ] = row
		}
		return arrays
	}

	public static Object[][] listsToArrays( List<List<Object>> lists ) {
		Object[][] arrays = new Object[ lists.size() ][]
		for ( int i in 0..( lists.size() - 1 ) ) {
			Object[] row
			if ( lists[ i ] instanceof List ) {
				row = lists[ i ].toArray()
			} else {
				row = new Object[ 1 ]
				row[ 0 ] = lists[ i ]
			}
			arrays[ i ] = row
		}
		return arrays
	}

	public static Object[][] valueToArrays( double value, int width, int height ) {
		dataToArrays( value, width, height )
	}

	public static Object[][] valueToArrays( double value, XCellRange range ) {
		dataToArrays( value, range )
	}

	public static Object[][] stringToArrays( String string, int width, int height ) {
		dataToArrays( string, width, height )
	}

	public static Object[][] stringToArrays( String string, XCellRange range ) {
		dataToArrays( string, range )
	}

	private static Object[][] dataToArrays( Object data, XCellRange range ) {
		CellRangeAddress address = range.address
		dataToArrays( data, address.EndColumn - address.StartColumn + 1, address.EndRow - address.StartRow + 1 )
	}

	private static Object[][] dataToArrays( Object data, int width, int height ) {
		Object[][] arrays = new Object[ height ][]
		for ( int i in 0..(arrays.length - 1) ) {
			Object[] array = new Object[ width ]
			for ( int j in 0..(array.length - 1) ) {
				array[ j ] = data
			}
			arrays[ i ] = array
		}
		arrays
	}

	public static List<List<Double>> doubleArraysToLists( Double[][] data ) {
		genericArraysToLists( data )
	}

	public static List<List<String>> stringArraysToLists( String[][] data ) {
		genericArraysToLists( data )
	}

	public static List<List<Object>> arraysToLists( Object[][] data ) {
		genericArraysToLists( data )
	}

	private static <T> List<List<T>> genericArraysToLists( T[][] data ) {
		List<List<T>> lists = new ArrayList<List<T>>( data.length )
		for ( T[] row in data ) {
			lists.add( row.toList() )
		}
		lists
	}

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
