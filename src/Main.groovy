import com.serwylo.gruno.spreadsheet.SpreadsheetConnector

SpreadsheetConnector connector = new SpreadsheetConnector()
connector.connect()

def docCsv = connector.open( "file:///tmp/test.csv" )

def doc1 = connector.open()
def sheet1 = doc1.sheets[ 0 ]
sheet1[ "A1" ]    << "A1"
sheet1[ "B1" ]    << "B1"
sheet1[ "C1:F1" ] << "C1->F1"

def doc2 = connector.open()
def sheet2 = doc2.sheets[ 0 ]
sheet2[ "A1" ]    << "1"
sheet2[ "B1" ]    << "2"
sheet2[ "C1:F1" ] << "3"

sheet2[ "A2" ]    << sheet1[ "A1" ]
sheet2[ "B2" ]    << sheet1[ "B1" ]
sheet2[ "C2:F2" ] << sheet1[ "C1:F1" ]

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