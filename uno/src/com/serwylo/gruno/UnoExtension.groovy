package com.serwylo.gruno

import com.sun.star.beans.XPropertySet
import com.sun.star.container.XEnumerationAccess
import com.sun.star.container.XIndexAccess
import com.sun.star.frame.XComponentLoader
import com.sun.star.lang.XComponent
import com.sun.star.lang.XMultiServiceFactory
import com.sun.star.sheet.XCellAddressable
import com.sun.star.sheet.XCellRangeAddressable
import com.sun.star.sheet.XCellRangeData
import com.sun.star.sheet.XSheetAnnotationAnchor
import com.sun.star.sheet.XSheetAnnotationsSupplier
import com.sun.star.sheet.XSheetOperation
import com.sun.star.sheet.XSpreadsheet
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.sheet.XSpreadsheets
import com.sun.star.table.CellAddress
import com.sun.star.table.XCell
import com.sun.star.table.XCellRange
import com.sun.star.table.XColumnRowRange
import com.sun.star.table.XTableColumns
import com.sun.star.table.XTableRows
import com.sun.star.text.XText
import com.sun.star.uno.UnoRuntime
import com.sun.star.beans.PropertyValue
import com.sun.star.util.XMergeable
import com.sun.star.util.XReplaceDescriptor
import com.sun.star.util.XReplaceable

class UnoExtension {

	public static XSpreadsheetDocument getDocument( XComponent component ) {
		UnoRuntime.queryInterface( XSpreadsheetDocument.class, component )
	}

	public static XComponent loadSpreadsheet( XComponentLoader loader ) {
		loader.loadComponentFromURL( "private:factory/scalc", "_blank", 0, new PropertyValue[0] );
	}

	public static XSpreadsheet getAt( XSpreadsheets spreadsheets, int index ) {
		XIndexAccess xSheetsIA = UnoRuntime.queryInterface( XIndexAccess.class, spreadsheets )
		UnoRuntime.queryInterface( XSpreadsheet.class, xSheetsIA.getByIndex( index ) )
	}

	public static XSpreadsheet getAt( XSpreadsheets spreadsheets, String key ) {
		UnoRuntime.queryInterface( XSpreadsheet.class, spreadsheets.getByName( key ) )
	}

	public static XText getText( XCell cell ) {
		UnoRuntime.queryInterface( XText.class, cell )
	}

	public static XEnumerationAccess getEnumerationAccess( XCell cell ) {
		UnoRuntime.queryInterface( XEnumerationAccess.class, cell );
	}

	public static XPropertySet getPropertySet( XCell cell ) {
		UnoRuntime.queryInterface( XPropertySet.class, cell )
	}

	public static void putAt( XPropertySet propertySet, String key, Object value ) {
		propertySet.setPropertyValue( key, value )
	}

	public static void getAt( XPropertySet propertySet, String key ) {
		propertySet.getPropertyValue( key )
	}

	public static XCellAddressable getAddressable( XCell cell ) {
		UnoRuntime.queryInterface( XCellAddressable.class, cell )
	}

	public static CellAddress getAddress( XCell cell ) {
		cell.addressable.cellAddress
	}

	public static XSheetAnnotationsSupplier getAnnotationSupplier( XSpreadsheet sheet ) {
		UnoRuntime.queryInterface( XSheetAnnotationsSupplier.class, sheet );
	}

	public static XSheetAnnotationAnchor getSheetAnnotationAnchor( XCell cell ) {
		UnoRuntime.queryInterface( XSheetAnnotationAnchor.class, cell )
	}

	public static XPropertySet getPropertySet( XCellRange range ) {
		UnoRuntime.queryInterface( XPropertySet.class, range )
	}

	public static XReplaceable getReplaceable( XCellRange range ) {
		UnoRuntime.queryInterface( XReplaceable.class, range );
	}

	public static XMergeable getMergeable( XCellRange cellRange ) {
        UnoRuntime.queryInterface( XMergeable.class, cellRange )
	}

	public static XColumnRowRange getColumnRowRange( XCellRange cellRange ) {
        UnoRuntime.queryInterface( XColumnRowRange.class, cellRange )
	}

	public static XTableColumns getColumns( XCellRange cellRange ) {
        cellRange.columnRowRange.columns
	}

	public static XTableRows getRows( XCellRange cellRange ) {
        cellRange.columnRowRange.rows
	}

	public static XCellRangeAddressable getCellRangeAddressable( XCellRange cellRange ) {
        UnoRuntime.queryInterface( XCellRangeAddressable.class, cellRange )
	}

	public static Object getAt( XIndexAccess access, Integer index ) {
		access.getByIndex( index )
	}

	public static XSheetOperation getSheetOperation( XCellRangeData data ) {
        UnoRuntime.queryInterface( XSheetOperation.class, data );
	}

	public static void setSearchWords( XReplaceDescriptor replaceDescriptor, boolean value ) {
        replaceDescriptor.setPropertyValue( "SearchWords", new Boolean( value ) );
	}

}


class UnoStaticExtension {

}