import com.serwylo.gruno.spreadsheet.SpreadsheetConnector
import com.serwylo.gruno.utils.Introspector
import com.sun.star.sheet.XCellRangeData
import com.sun.star.table.XCellRange
import com.sun.star.beans.XIntrospection;

def spreadsheetProgram = new SpreadsheetConnector()

def doc1 = spreadsheetProgram.open()

def sheet = doc1.sheets[ 0 ]

def size = 250000;
def numbers = 1..size

print "Adding random values... "
[ 'A', 'B' ].each {
	List<List<Object>> data = numbers.collect { [ Math.random() * 100 ] }
	sheet[ "${it}1:$it$size" ] << data
}
println "done!"

println "Adding const values... "
sheet[ "C1:C$size" ] << 3
println "done!"

println "Adding const strings... "
sheet[ "D1:D$size" ] << "Total:"
println "done!"

print "Adding sums... "
sheet[ "E1:E$size" ].formulas = numbers.collect { [ "=SUM(A$it:C$it" ] }
println "Done!"

/*
print "Adding more random values... "
sheet[ "B1:B200000" ] << Math.random() * 10
println "Done!"

print "Summing... "
for ( int i = 0; i < 200000; i ++ ) {
	sheet[ "C$i" ].formula = "=SUM(A$i:B$i)"
	if ( i % 10000 ) {
		println " $i"
	}
}
println "Done!"

sleep 1000
*/

System.exit( 0 )

/*
try
{
	GroovySpreadsheetSample aSample = new GroovySpreadsheetSample( args );
	aSample.doSampleFunction();
}
catch (Exception ex)
{
	println( "Error: Sample caught exception!\nException Message = $ex.message" );
	ex.printStackTrace();
	System.exit( 1 );
}
System.out.println( "\nSamples done." );
System.exit( 0 );
*/