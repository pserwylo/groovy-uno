import com.serwylo.gruno.spreadsheet.SpreadsheetConnector

SpreadsheetConnector connector = new SpreadsheetConnector()
connector.connect()

def doc1 = connector.open()
def sheet1 = doc1.sheets[ 0 ]

long time = System.currentTimeMillis()

for ( def i = 0; i < 1000; i ++ ) {
	sheet1[ "A1" ] << i
}

println System.currentTimeMillis() - time

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