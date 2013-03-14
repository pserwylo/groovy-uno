groovy-uno
==========

(Hopefully) makes scripting OpenOffice/LibreOffice easier.

Although I'm aware that there are several language bindings for UNO (e.g. C++/Java/OpenOffice Basic), 
they are all very terse and I find annoying to work with. Even the Basic scripting language used for
scripting within OpenOffice is a little annoying (at least for me).

I had a need to do some data munging for some friends, and I have also been using Groovy for an unrelated
project recently. Because it is a JVM language, it can make use of the Java UNO bindings. Also, because 
it supports metaClasses, I can dynamically add methods and use appropriate operator overloading to make 
the whole proccess of working with UNO much nicer.

Coverage
========

At this stage, I'm just adding features as required, rather than providing a complete implementation 
of every property feature defined in the IDL. Think of this as more a utility library, rather than
a Groovy implementation of the IDL. As I require more features, I'll add them either as methods on
the existing Java UNO classes, or as utility classes.

Examples
========

Connecting to running office:
---------

Java UNO:

XComponentContext context           = Bootstrap.bootstrap();
XMultiServiceFactory serviceFactory = context.getServiceManager();
XComponentContext context           = uno.util.Bootstrap.bootstrap();
XComponentLoader loader             = (XComponentLoader) UnoRuntime.queryInterface( XComponentLoader.class, serviceFactory.createInstanceWithContext( "com.sun.star.frame.Desktop", context ) );
XComponent component                = loader.loadComponentFromURL( "private:factory/scalc", "_blank", 0, new PropertyValue[0] );
XSpreadsheetDocument document       = (XSpreadsheetDocument) UnoRuntime.queryInterface( XSpreadsheetDocument.class, component );

groovy-uno:

SpreadsheetConnector connector = new SpreadsheetConnector()
XSpreadsheetDocument document  = connector.open()

Get spreadsheet at index:
---------

Java UNO:

XSpreadsheets sheets     = document.getSheets();
XIndexAccess indexAccess = (XIndexAccess) UnoRuntime.queryInterface( XIndexAccess.class, sheets );
XSpreadsheet sheet       = (XSpreadsheet) UnoRuntime.queryInterface( XSpreadsheet.class, indexAccess.getByIndex( index ) );

groovy-uno:

XSpreadsheet sheet = document[ index ]

Setting value at position:
-------

Java UNO:

XSpreadsheet sheet = ...;
sheet.getCellRangeByName( cellName ).getCellByPosition( 0, 0 ).setValue( value );

groovy-uno:

XSpreadsheet sheet = ...
sheet[ cellName ] << value
