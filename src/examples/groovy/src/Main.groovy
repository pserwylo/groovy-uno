import com.serwylo.uno.spreadsheet.SpreadsheetConnector

def spreadsheetProgram = new SpreadsheetConnector()
def newDoc = spreadsheetProgram.open()
def sheet = newDoc.sheets[ 0 ]

def size = 250000;
def numbers = 1..size

print "Adding random values... "
'ABCDEFGHIJK'.each {
	List<Object> data = numbers.collect { Math.random() * 100 }
	sheet[ "${it}1:$it$size" ] << data
}
println "done!"

print "Adding sums... "
sheet[ "L1:L$size" ].formulas = numbers.collect { "=SUM(A$it:C$it" }
println "Done!"

System.exit( 0 )
