package com.serwylo.uno.extension

import com.serwylo.uno.utils.DataUtils
import com.sun.star.beans.XPropertySet
import com.sun.star.container.XEnumerationAccess
import com.sun.star.container.XIndexAccess
import com.sun.star.lang.XComponent
import com.sun.star.sdbc.XCloseable
import com.sun.star.sheet.XArrayFormulaRange
import com.sun.star.sheet.XCellAddressable
import com.sun.star.sheet.XCellRangeAddressable
import com.sun.star.sheet.XCellRangeData
import com.sun.star.sheet.XCellRangeFormula
import com.sun.star.sheet.XSheetAnnotationAnchor
import com.sun.star.sheet.XSheetAnnotationsSupplier
import com.sun.star.sheet.XSheetOperation
import com.sun.star.sheet.XSpreadsheet
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.sheet.XSpreadsheets
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



	// ========= XSpreadsheets =========

	public static XSpreadsheet getAt( XSpreadsheets self, int index ) {
		XIndexAccess xSheetsIA = cache( self, "indexAccess$index" ) {
			UnoRuntime.queryInterface( XIndexAccess.class, self )
		}

		cache( self, "getAt$index" ) {
			UnoRuntime.queryInterface( XSpreadsheet.class, xSheetsIA.getByIndex( index ) )
		}
	}

	public static XSpreadsheet getAt( XSpreadsheets self, String key ) {
		cache( self, "getAt$key" ) {
			UnoRuntime.queryInterface( XSpreadsheet.class, self.getByName( key ) )
		}
	}



	// ========= XSpreadsheet =========

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
		cache( self, "getForumlaRange" ) {
			UnoRuntime.queryInterface( XCellRangeFormula.class, self )
		}
	}

	public static void setFormulas( XCellRange self, List<List<String>> formulas ) {
		XCellRangeFormula rangeFormulas = self.cellRangeFormula
		rangeFormulas.formulaArray = DataUtils.listsToArrays( formulas )
	}

	public static XCellRangeData getCellRangeData( XCellRange self ) {
		cache( self, "getDataRange" ) {
			UnoRuntime.queryInterface( XCellRangeData.class, self )
		}
	}

	public static void setData( XCellRange self, Object[][] data ) {
		XCellRangeData rangeData = self.cellRangeData
		rangeData.dataArray = data
	}

	public static void leftShift( XCellRange self, List<List<Object>> data ) {
		self.data = DataUtils.listsToArrays( data )
	}

	public static void leftShift( XCellRange self, double value ) {
		self.data = DataUtils.valueToArrays( value, self )
	}

	public static void leftShift( XCellRange self, String string ) {
		self.data = DataUtils.stringToArrays( string, self )
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

	public static XCellRangeData getData( XCellRange self ) {
        cache( self, "getData", ) {
			UnoRuntime.queryInterface( XCellRangeData.class, self )
		}
	}

	public static XArrayFormulaRange getFormulaRange( XCellRange self ) {
        cache( self, "getFormulaRange", ) {
			UnoRuntime.queryInterface( XArrayFormulaRange.class, self )
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

	public static void getAt( XPropertySet self, String key ) {
		self.getPropertyValue( key )
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
