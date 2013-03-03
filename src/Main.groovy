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