/*************************************************************************
 *
 *  The Contents of this file are made available subject to the terms of
 *  the BSD license.
 *
 *  Copyright 2000, 2010 Oracle and/or its affiliates.
 *  All rights reserved.
 *
 *  Redistribution and use in source and binary forms, with or without
 *  modification, are permitted provided that the following conditions
 *  are met:
 *  1. Redistributions of source code must retain the above copyright
 *     notice, this list of conditions and the following disclaimer.
 *  2. Redistributions in binary form must reproduce the above copyright
 *     notice, this list of conditions and the following disclaimer in the
 *     documentation and/or other materials provided with the distribution.
 *  3. Neither the name of Sun Microsystems, Inc. nor the names of its
 *     contributors may be used to endorse or promote products derived
 *     from this software without specific prior written permission.
 *
 *  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
 *  "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
 *  LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
 *  FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
 *  COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
 *  INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
 *  BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS
 *  OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 *  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR
 *  TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE
 *  USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *
 *************************************************************************/



import com.serwylo.uno.Connector
import com.serwylo.uno.spreadsheet.SpreadsheetConnector
import com.serwylo.uno.utils.ColumnUtils
import com.sun.star.lang.XMultiComponentFactory
import com.sun.star.sheet.XSpreadsheet
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.sheet.XSpreadsheets
import com.sun.star.table.CellAddress
import com.sun.star.table.CellRangeAddress
import com.sun.star.uno.RuntimeException;
import com.sun.star.uno.UnoRuntime
import com.sun.star.uno.XComponentContext;

/** This is a helper class for the spreadsheet and table samples.
    It connects to a running office and creates a spreadsheet document.
    Additionally it contains various helper functions.
 */
public class GroovySpreadsheetDocHelper
{

	private SpreadsheetConnector connector

    private XSpreadsheetDocument document;

    public GroovySpreadsheetDocHelper( String[] args )
    {
		connector = new SpreadsheetConnector()
       	document = connector.open()
    }

// __  helper methods  ____________________________________________

	public SpreadsheetConnector getConnector()
	{
		connector
	}

    /** Inserts a new empty spreadsheet with the specified name.
        @param aName  The name of the new sheet.
        @param nIndex  The insertion index.
        @return  The XSpreadsheet interface of the new sheet. */
    public XSpreadsheet insertSpreadsheet( String aName, short nIndex )
    {
    	document.sheets.insertNewByName( aName, nIndex )
        document.sheets[ aName ]
    }

// ________________________________________________________________
// Methods to fill values into cells.

    /** Writes a double value into a spreadsheet.
        @param xSheet  The XSpreadsheet interface of the spreadsheet.
        @param aCellName  The address of the cell (or a named range).
        @param fValue  The value to write into the cell. */
    public void setValue(
		XSpreadsheet xSheet,
		String aCellName,
		double fValue ) throws RuntimeException, Exception
    {
        xSheet.getCellAt( aCellName ) << fValue;
    }

    /** Writes a formula into a spreadsheet.
        @param xSheet  The XSpreadsheet interface of the spreadsheet.
        @param aCellName  The address of the cell (or a named range).
        @param aFormula  The formula to write into the cell. */
    public void setFormula(
            com.sun.star.sheet.XSpreadsheet xSheet,
            String aCellName,
            String aFormula ) throws RuntimeException, Exception
    {
        xSheet.getCellAt( aCellName ) << aFormula
    }

    /** Writes a date with standard date format into a spreadsheet.
        @param xSheet  The XSpreadsheet interface of the spreadsheet.
        @param aCellName  The address of the cell (or a named range).
        @param nDay  The day of the date.
        @param nMonth  The month of the date.
        @param nYear  The year of the date. */
    public void setDate(
            com.sun.star.sheet.XSpreadsheet xSheet,
            String aCellName,
            int nDay, int nMonth, int nYear ) throws RuntimeException, Exception
    {
        // Set the date value.
        com.sun.star.table.XCell xCell = xSheet.getCellRangeByName( aCellName ).getCellByPosition( 0, 0 );
        String aDateStr = nMonth + "/" + nDay + "/" + nYear;
        xCell.setFormula( aDateStr );

        // Set standard date format.
        com.sun.star.util.XNumberFormatsSupplier xFormatsSupplier =
            (com.sun.star.util.XNumberFormatsSupplier) UnoRuntime.queryInterface(
                com.sun.star.util.XNumberFormatsSupplier.class, getDocument() );
        com.sun.star.util.XNumberFormatTypes xFormatTypes =
            (com.sun.star.util.XNumberFormatTypes) UnoRuntime.queryInterface(
                com.sun.star.util.XNumberFormatTypes.class, xFormatsSupplier.getNumberFormats() );
        int nFormat = xFormatTypes.getStandardFormat(
            com.sun.star.util.NumberFormat.DATE, new com.sun.star.lang.Locale() );

        com.sun.star.beans.XPropertySet xPropSet =
            UnoRuntime.queryInterface( com.sun.star.beans.XPropertySet.class, xCell );
        xPropSet.setPropertyValue( "NumberFormat", new Integer( nFormat ) );
    }

    /** Draws a colored border around the range and writes the headline in the
        first cell.
        @param xSheet  The XSpreadsheet interface of the spreadsheet.
        @param aRange  The address of the cell range (or a named range).
        @param aHeadline  The headline text. */
    public void prepareRange(
            com.sun.star.sheet.XSpreadsheet xSheet,
            String aRange, String aHeadline ) throws RuntimeException, Exception
    {
        com.sun.star.beans.XPropertySet xPropSet = null;
        com.sun.star.table.XCellRange xCellRange = null;

        // draw border
        xCellRange = xSheet.getCellRangeByName( aRange );
        xPropSet = UnoRuntime.queryInterface( com.sun.star.beans.XPropertySet.class, xCellRange );
        com.sun.star.table.BorderLine aLine = new com.sun.star.table.BorderLine();
        aLine.Color = 0x99CCFF;
        aLine.InnerLineWidth = aLine.LineDistance = 0;
        aLine.OuterLineWidth = 100;
        com.sun.star.table.TableBorder aBorder = new com.sun.star.table.TableBorder();
        aBorder.TopLine = aBorder.BottomLine = aBorder.LeftLine = aBorder.RightLine = aLine;
        aBorder.IsTopLineValid = aBorder.IsBottomLineValid = true;
        aBorder.IsLeftLineValid = aBorder.IsRightLineValid = true;
        xPropSet.setPropertyValue( "TableBorder", aBorder );

        // draw headline
        com.sun.star.sheet.XCellRangeAddressable xAddr =
            UnoRuntime.queryInterface( com.sun.star.sheet.XCellRangeAddressable.class, xCellRange );
        com.sun.star.table.CellRangeAddress aAddr = xAddr.getRangeAddress();

        xCellRange = xSheet.getCellRangeByPosition(
            aAddr.StartColumn, aAddr.StartRow, aAddr.EndColumn, aAddr.StartRow );
        xPropSet =
            UnoRuntime.queryInterface( com.sun.star.beans.XPropertySet.class, xCellRange );
        xPropSet.setPropertyValue( "CellBackColor", new Integer( 0x99CCFF ) );
        // write headline
        com.sun.star.table.XCell xCell = xCellRange.getCellByPosition( 0, 0 );
        xCell.setFormula( aHeadline );
        xPropSet =
            UnoRuntime.queryInterface( com.sun.star.beans.XPropertySet.class, xCell );
        xPropSet.setPropertyValue( "CharColor", new Integer( 0x003399 ) );
        xPropSet.setPropertyValue( "CharWeight", new Float( com.sun.star.awt.FontWeight.BOLD ) );
    }

// ________________________________________________________________
// Methods to create cell addresses and range addresses.

    /** Creates a com.sun.star.table.CellAddress and initializes it
        with the given range.
        @param xSheet  The XSpreadsheet interface of the spreadsheet.
        @param aCell  The address of the cell (or a named cell). */
    public CellAddress createCellAddress( XSpreadsheet xSheet, String aCell ) throws RuntimeException, Exception {
        xSheet.getCellAt( aCell ).address;
    }

    /** Creates a com.sun.star.table.CellRangeAddress and initializes
        it with the given range.
        @param xSheet  The XSpreadsheet interface of the spreadsheet.
        @param aRange  The address of the cell range (or a named range). */
    public CellRangeAddress createCellRangeAddress( XSpreadsheet xSheet, String aRange ) {
        xSheet[ aRange ].address
    }

// ________________________________________________________________
// Methods to convert cell addresses and range addresses to strings.

    /** Returns the text address of the cell.
        @param nColumn  The column index.
        @param nRow  The row index.
        @return  A string containing the cell address. */
    public String getCellAddressString( int nColumn, int nRow )
    {
        String columnName = ColumnUtils.indexToName( nColumn )
		"$columnName$nRow"
    }

    /** Returns the text address of the cell range.
        @param aCellRange  The cell range address.
        @return  A string containing the cell range address. */
    public String getCellRangeAddressString(
            com.sun.star.table.CellRangeAddress aCellRange )
    {
		def start = getCellAddressString( aCellRange.StartColumn, aCellRange.StartRow )
		def end   = getCellAddressString( aCellRange.EndColumn, aCellRange.EndRow );
        return "$start:$end"
    }

    /** Returns the text address of the cell range.
        @param xCellRange  The XSheetCellRange interface of the cell range.
        @param bWithSheet  true = Include sheet name.
        @return  A string containing the cell range address. */
    public String getCellRangeAddressString(
            com.sun.star.sheet.XSheetCellRange xCellRange,
            boolean bWithSheet )
    {
        String aStr = "";
        if (bWithSheet)
        {
            com.sun.star.sheet.XSpreadsheet xSheet = xCellRange.getSpreadsheet();
            com.sun.star.container.XNamed xNamed =
                UnoRuntime.queryInterface( com.sun.star.container.XNamed.class, xSheet );
            aStr += xNamed.getName() + ".";
        }
        com.sun.star.sheet.XCellRangeAddressable xAddr =
            UnoRuntime.queryInterface( com.sun.star.sheet.XCellRangeAddressable.class, xCellRange );
        aStr += getCellRangeAddressString( xAddr.getRangeAddress() );
        return aStr;
    }

    /** Returns a list of addresses of all cell ranges contained in the collection.
        @param xRangesIA  The XIndexAccess interface of the collection.
        @return  A string containing the cell range address list. */
    public String getCellRangeListString(
            com.sun.star.container.XIndexAccess xRangesIA ) throws RuntimeException, Exception
    {
        String aStr = "";
        int nCount = xRangesIA.getCount();
        for (int nIndex = 0; nIndex < nCount; ++nIndex)
        {
            if (nIndex > 0)
                aStr += " ";
            Object aRangeObj = xRangesIA.getByIndex( nIndex );
            com.sun.star.sheet.XSheetCellRange xCellRange =
                UnoRuntime.queryInterface( com.sun.star.sheet.XSheetCellRange.class, aRangeObj );
            aStr += getCellRangeAddressString( xCellRange, false );
        }
        return aStr;
    }

// ________________________________________________________________
}