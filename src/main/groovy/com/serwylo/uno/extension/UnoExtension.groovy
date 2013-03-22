package com.serwylo.uno.extension

import com.serwylo.uno.utils.ColumnUtils
import com.serwylo.uno.utils.DataUtils
import com.sun.star.beans.XPropertySet
import com.sun.star.container.XEnumerationAccess
import com.sun.star.container.XIndexAccess
import com.sun.star.container.XNameAccess
import com.sun.star.container.XNamed
import com.sun.star.lang.XComponent
import com.sun.star.sdbc.XCloseable
import com.sun.star.sheet.FillDirection
import com.sun.star.sheet.XArrayFormulaRange
import com.sun.star.sheet.XCellAddressable
import com.sun.star.sheet.XCellRangeAddressable
import com.sun.star.sheet.XCellRangeData
import com.sun.star.sheet.XCellRangeFormula
import com.sun.star.sheet.XCellSeries
import com.sun.star.sheet.XSheetAnnotationAnchor
import com.sun.star.sheet.XSheetAnnotationsSupplier
import com.sun.star.sheet.XSheetCellCursor
import com.sun.star.sheet.XSheetOperation
import com.sun.star.sheet.XSpreadsheet
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.sheet.XSpreadsheets
import com.sun.star.sheet.XUsedAreaCursor
import com.sun.star.table.CellAddress
import com.sun.star.table.CellRangeAddress
import com.sun.star.table.XCell
import com.sun.star.table.XCellRange
import com.sun.star.table.XColumnRowRange
import com.sun.star.table.XTableColumns
import com.sun.star.table.XTableRows
import com.sun.star.text.XText
import com.sun.star.uno.UnoRuntime
import com.sun.star.util.CloseVetoException
import com.sun.star.util.XMergeable
import com.sun.star.util.XReplaceDescriptor
import com.sun.star.util.XReplaceable

import java.awt.Point
import java.awt.geom.Dimension2D
import java.awt.geom.Point2D

class UnoExtension {

	// http://groovy.codehaus.org/ExpandoMetaClass+-+Properties
	private static properties = Collections.synchronizedMap([:])

	private static Closure  cache = { Object self, String propName, Closure<Object> closure ->
		Object value
		String key = System.identityHashCode( self ) + "-" + propName
		if ( properties.containsKey( key ) ) {
			value = properties[ key ]
		} else {
			value = closure()
			properties[ key ] = value
		}
		return value
	}


	private static XComponent spreadsheetDocumentToComponent( XSpreadsheetDocument doc ) {
		cache( doc, "spreadsheetComponent" ) {
			UnoRuntime.queryInterface( XComponent.class, doc )
		}
	}


	// ========= XSpreadsheetDocument =========

	public static boolean close( XSpreadsheetDocument self ) {
		XCloseable closeable = UnoRuntime.queryInterface( XCloseable.class, self )
		boolean success = false
		if ( closeable ) {
			try {
				closeable.close()
				success = true
			} catch ( CloseVetoException e ) {
				success = false
			}
		} else {
			XComponent component = spreadsheetDocumentToComponent( self )
			component.dispose()
			success = true
		}
		return success
	}


	public static XSpreadsheet getAt( XSpreadsheetDocument self, int index ) {
		self.getSheets().getAt( index )
	}


	public static XSpreadsheet getAt( XSpreadsheetDocument self, String sheetName ) {
		self.getSheets().getAt( sheetName )
	}


	// ========= XSpreadsheets =========

	public static XIndexAccess getIndexAccess( XSpreadsheets self ) {
		cache( self, "indexAccess" ) {
			UnoRuntime.queryInterface( XIndexAccess.class, self )
		}
	}

	public static int size( XSpreadsheets self ) {
		self.getIndexAccess().size()
	}

	public static XSpreadsheet getAt( XSpreadsheets self, int index ) {
		cache( self, "getAt$index" ) {
			UnoRuntime.queryInterface( XSpreadsheet.class, self.getIndexAccess().getByIndex( index ) )
		}
	}



	// ========= XSpreadsheet =========

	public static CellAddress getDimensions( XSpreadsheet self ) {

		XSheetCellCursor xSheetCellCursor = self.createCursor()
		XUsedAreaCursor xUsedAreaCursor = UnoRuntime.queryInterface( XUsedAreaCursor.class, xSheetCellCursor )
		xUsedAreaCursor.gotoEndOfUsedArea( false )

		xSheetCellCursor.getCellByPosition( 0, 0 ).address
	}

	public static XNamed toNamed( XSpreadsheet self ) {
		cache( self, "toNamed" ) {
			UnoRuntime.queryInterface( XNamed.class, self )
		}
	}

	public static String getName( XSpreadsheet self ) {
		self.toNamed().getName()
	}

	public static String setName( XSpreadsheet self, String value ) {
		self.toNamed().setName( value )
	}

	public static XCellRange toCellRange( XSpreadsheet self ) {
		cache( self, "toCellRange" ) {
			UnoRuntime.queryInterface( XCellRange.class, self )
		}
	}

	public static XCellRange getAt( XSpreadsheet self, String key ) {
		cache( self, "getAt$key" ) {
			UnoRuntime.queryInterface( XCellRange.class, self.getCellRangeByName( key ) )
		}
	}

	public static XCellRange getCells( XSpreadsheet self ) {
		cache( self, "getCells" ) {
			UnoRuntime.queryInterface( XCellRange.class, self )
		}
	}

	public static XCell getCellAt( XSpreadsheet self, String key ) {
		cache( self, "getCellAt$key" ) {
			if ( key.contains( ':' ) ) {
				throw new IllegalArgumentException( "Cannot specify a cell range, must be a single cell." )
			}
			self[ key ].getCellByPosition( 0, 0 )
		}
	}

	public static XCell getAt( XSpreadsheet self, int column, int row ) {
		cache( self, "getAt$column,$row" ) {
			self.getCellByPosition( column, row )
		}
	}

	public static XCellRange getAt( XSpreadsheet self, int column ) {
		cache( self, "getAt$column" ) {
			self.toCellRange().columns[ column ]
		}
	}

	public static XSheetAnnotationsSupplier getAnnotationSupplier( XSpreadsheet self ) {
		cache( self, "getAnnotationSupplier" ) {
			UnoRuntime.queryInterface( XSheetAnnotationsSupplier.class, self );
		}
	}

	public static void clear( XCell self ) {
		self << ""
	}



	// ========= XCellRange =========

	public static XCellSeries toCellSeries( XCellRange self ) {
		cache( self, "toCellSeries" ) {
			UnoRuntime.queryInterface( XCellSeries.class, self )
		}
	}


	public static void fillDown( XCellRange self ) {
		self.fill( FillDirection.TO_BOTTOM )
	}


	public static void fill( XCellRange self, FillDirection direction ) {
		XCellSeries series = self.toCellSeries()
		series.fillAuto( direction, 1 )
	}

	/**
	 * Iterates over each cell in the range (rows first, then columns).
	 * @param self
	 * @param closure Takes an XCell as its only parameter.
	 */
	public static void each( XCellRange self, Closure closure ) {
		CellRangeAddress address = self.address
		for ( int column = address.StartColumn; column <= address.EndColumn; column ++ ) {
			for ( int row = address.StartRow; row <= address.EndRow; row ++ ) {
				XCell cell = self.getCellByPosition( column - address.StartColumn, row - address.StartRow )
				closure( cell )
			}
		}
	}

	public static void clear( XCellRange self ) {
		self.operation.clearContents( 1 | 2 | 4 | 8 | 16 | 32 | 64 | 128 | 256 | 512 | 1024 | 2048 | 4096 )
	}

	public static void setValue( XCellRange self, double value ) {
		self.getCellByPosition( 0, 0 ).setValue( value )
	}

	public static double getValue( XCellRange self ) {
		self.getCellByPosition( 0, 0 ).getValue()
	}

	public static void setFormula( XCellRange self, String value ) {
		self.getCellByPosition( 0, 0 ).setFormula( value )
	}

	public static String getFormula( XCellRange self ) {
		self.getCellByPosition( 0, 0 ).getFormula()
	}

	public static XCellRangeFormula getCellRangeFormula( XCellRange self ) {
		UnoRuntime.queryInterface( XCellRangeFormula.class, self )
	}

	public static List<List<String>> getFormulas( XCellRange self ) {
		DataUtils.stringArraysToLists( self.getCellRangeFormula().getFormulaArray() )
	}

	public static void setFormulas( XCellRange self, List<List<String>> formulas ) {
		XCellRangeFormula rangeFormulas = self.getCellRangeFormula()
		rangeFormulas.setFormulaArray( DataUtils.stringListsToArrays( formulas ) )
	}

	/**
	 * Unfotunately if I try to overload setFormulas( self, String[][] formulas ) then it still invokes
	 * the List<List<String>> method. As such, I've put this method here for those times when we have hundreds of
	 * thousands of rows, and it just seems silly to convert from Lists to arrays.
	 * @param self
	 * @param formulas
	 */
	public static void setFormulasFromArrays( XCellRange self, String[][] formulas ) {
		XCellRangeFormula rangeFormulas = self.getCellRangeFormula()
		rangeFormulas.setFormulaArray( formulas )
	}

	public static XCellRangeData getCellRangeData( XCellRange self ) {
		cache( self, "getDataRange" ) {
			UnoRuntime.queryInterface( XCellRangeData.class, self )
		}
	}

	public static void setData( XCellRange self, Object[][] data ) {
		XCellRangeData rangeData = self.getCellRangeData()
		rangeData.setDataArray( data )
	}

	public static void setData( XCellRange self, List<List<Object>> data ) {
		self.setData( DataUtils.listsToArrays( data ) )
	}

	public static List<List<Object>> getData( XCellRange self ) {
		DataUtils.arraysToLists( self.getCellRangeData().getDataArray() )
	}

	public static void leftShift( XCellRange self, List<List<Object>> data ) {
		self.setData( data )
	}

	public static void leftShift( XCellRange self, double value ) {
		self.setData( DataUtils.valueToArrays( value, self ) )
	}

	public static void leftShift( XCellRange self, String string ) {
		self.setData( DataUtils.stringToArrays( string, self ) )
	}

	public static void leftShift( XCellRange self, XCellRange sourceRange ) {

		CellRangeAddress destAddress   = self.address
		CellRangeAddress sourceAddress = sourceRange.address

		int destColumns   = destAddress.EndColumn   - destAddress.StartColumn
		int destRows      = destAddress.EndRow      - destAddress.StartRow
		int sourceColumns = sourceAddress.EndColumn - sourceAddress.StartColumn
		int sourceRows    = sourceAddress.EndRow    - sourceAddress.StartRow

		if ( destColumns != sourceColumns || destRows != sourceRows ) {
			throw new IllegalArgumentException( "Both cell ranges must be the same size." )
		}

		for ( int column = 0; column <= destColumns; column ++ ) {
			for ( int row = 0; row <= destRows; row ++ ) {
				XCell sourceCell = sourceRange.getCellByPosition( column, row )
				XCell destCell   = self.getCellByPosition( column, row )
				destCell << sourceCell
			}
		}
	}

	public static CellRangeAddress getAddress( XCellRange self ) {
		cache( self, "getAddress" ) {
			XCellRangeAddressable xAddr = UnoRuntime.queryInterface( XCellRangeAddressable.class, self )
			xAddr.rangeAddress
		}
	}

	public static XPropertySet getPropertySet( XCellRange self ) {
		cache( self, "getPropertySet" ) {
			UnoRuntime.queryInterface( XPropertySet.class, self )
		}
	}

	public static XReplaceable getReplaceable( XCellRange self ) {
		cache( self, "getReplaceable" ) {
			UnoRuntime.queryInterface( XReplaceable.class, self );
		}
	}

	public static XMergeable getMergeable( XCellRange self ) {
        cache( self, "getMergeable" ) {
			UnoRuntime.queryInterface( XMergeable.class, self )
		}
	}

	public static XColumnRowRange getColumnRowRange( XCellRange self ) {
        cache( self, "getColumnRowRange" ) {
			UnoRuntime.queryInterface( XColumnRowRange.class, self )
		}
	}

	public static XTableColumns getColumns( XCellRange self ) {
        self.columnRowRange.columns
	}

	public static XTableRows getRows( XCellRange self ) {
        self.columnRowRange.rows
	}

	public static XCellRangeAddressable getCellRangeAddressable( XCellRange self ) {
        cache( self, "getCellRangeAddressable" ) {
			UnoRuntime.queryInterface( XCellRangeAddressable.class, self )
		}
	}

	public static XSheetOperation getSheetOperation( XCellRange self ) {
        cache( self, "getSheetOperation" ) {
			UnoRuntime.queryInterface( XSheetOperation.class, self );
		}
	}

	public static XSheetOperation getOperation( XCellRange self ) {
        self.sheetOperation
	}

	public static XArrayFormulaRange getFormulaRange( XCellRange self ) {
        cache( self, "getFormulaRange", ) {
			UnoRuntime.queryInterface( XArrayFormulaRange.class, self )
		}
	}


	public static String getName( XCellRange self ) {
		cache( self, "getName" ) {
			CellRangeAddress address = self.getAddress()
			String startCol = ColumnUtils.indexToName( address.StartColumn )
			String endCol   = ColumnUtils.indexToName( address.EndColumn )
			int    startRow = address.StartRow + 1
			int    endRow   = address.EndRow + 1

			"$startCol$startRow:$endCol$endRow"
		}
	}


	// ========= XCell =========

	public static void leftShift( XCell self, XCell sourceCell ) {
		String formula = sourceCell.formula
		if ( formula ) {
			self.formula = formula
		} else {
			self.value = sourceCell.value
		}
	}

	public static void leftShift( XCell self, String formula ) {
		self.formula = formula
	}

	public static void leftShift( XCell self, double value ) {
		self.value = value
	}

	public static XText getText( XCell self ) {
		cache( self, "getText" ) {
			UnoRuntime.queryInterface( XText.class, self )
		}
	}

	public static XEnumerationAccess getEnumerationAccess( XCell self ) {
		cache( self, "getEnumerationAccess" ) {
			UnoRuntime.queryInterface( XEnumerationAccess.class, self );
		}
	}

	public static XEnumerationAccess( XCell self ) {
		self.getEnumerationAccess()
	}

	public static XPropertySet getPropertySet( XCell self ) {
		cache( self, "getPropertySet" ) {
			UnoRuntime.queryInterface( XPropertySet.class, self )
		}
	}

	public static XCellAddressable getAddressable( XCell self ) {
		cache( self, "getAddressable" ) {
			UnoRuntime.queryInterface( XCellAddressable.class, self )
		}
	}

	public static CellAddress getAddress( XCell self ) {
		self.addressable.cellAddress
	}

	public static XSheetAnnotationAnchor getSheetAnnotationAnchor( XCell self ) {
		cache( self, "getSheetAnnotationAnchor" ) {
			UnoRuntime.queryInterface( XSheetAnnotationAnchor.class, self )
		}
	}



	// ========= XPropertySet =========

	public static void putAt( XPropertySet self, String key, Object value ) {
		self.setPropertyValue( key, value )
	}

	public static Object getAt( XPropertySet self, String key ) {
		self.getPropertyValue( key )
	}



	// ========= XNameAccess =========

	public static Object getAt( XNameAccess self, String key ) {
		self.getByName( key )
	}

	public static boolean contains( XNameAccess self, String key ) {
		self.hasByName( key )
	}

	public static XSpreadsheet getAt( XSpreadsheets self, String key ) {
		XNameAccess nameAccess = self.toNameAccess()
		UnoRuntime.queryInterface( XSpreadsheet.class, nameAccess.getByName( key ) )
	}

	public static boolean contains( XSpreadsheets self, String key ) {
		XNameAccess nameAccess = self.toNameAccess()
		nameAccess.hasByName( key )
	}

	public static XNameAccess toNameAccess( XSpreadsheets self ) {
		cache( self, "toNameAccess" ) {
			UnoRuntime.queryInterface( XNameAccess.class, self )
		}
	}


	// ========= XIndexAccess =========

	public static Object getAt( XIndexAccess self, Integer index ) {
		self.getByIndex( index )
	}


	public static int size( XIndexAccess self ) {
		self.getCount()
	}

	/**
	 * More specific version, which returns an XCellRange.
	 * @param self
	 * @param index
	 * @return
	 */
	public static XCellRange getAt( XTableColumns self, Integer index ) {
		cache( self, "getAt$index" ) {
			UnoRuntime.queryInterface( XCellRange.class, self.getByIndex( index ) )
		}
	}

	/**
	 * More specific version, which returns an XCellRange.
	 * @param self
	 * @param index
	 * @return
	 */
	public static XCellRange getAt( XTableRows self, Integer index ) {
		cache( self, "getAt$index" ) {
			UnoRuntime.queryInterface( XCellRange.class, self.getByIndex( index ) )
		}
	}



	// ========= XTableRows/XTableColumns =========

	public static void removeAt( XTableRows self, Integer index, Integer numRows = 1 ) {
		self.removeByIndex( index, numRows )
	}

	public static void removeAt( XTableColumns self, Integer index, Integer numRows = 1 ) {
		self.removeByIndex( index, numRows )
	}



	// ========= XCellRangeData =========

	public static XSheetOperation getSheetOperation( XCellRangeData self ) {
        cache( self, "getSheetOperation" ) {
			UnoRuntime.queryInterface( XSheetOperation.class, self );
		}
	}



	// ========= XReplaceDescriptor =========

	public static void setSearchWords( XReplaceDescriptor self, boolean value ) {
        self.setPropertyValue( "SearchWords", new Boolean( value ) );
	}

}
