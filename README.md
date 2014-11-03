# groovy-uno

(Hopefully) makes scripting OpenOffice/LibreOffice easier.

Although I'm aware that there are several language bindings for UNO (e.g. C++/Java/OpenOffice Basic), 
they are all very terse and I find annoying to work with. Even the Basic scripting language used for
scripting within OpenOffice is a little annoying (at least for me).

I had a need to do some data munging for some friends, and I have also been using Groovy for an unrelated
project recently. Because it is a JVM language, it can make use of the Java UNO bindings. Also, because 
it supports metaClasses, I can dynamically add methods and use appropriate operator overloading to make 
the whole proccess of working with UNO much nicer.

## Coverage

At this stage, I'm just adding features as required, rather than providing a complete implementation 
of every property feature defined in the IDL. Think of this as more a utility library, rather than
a Groovy implementation of the IDL. As I require more features, I'll add them either as methods on
the existing Java UNO classes, or as utility classes.

## Building

Run `gradle build` from the top directory of this repository. It will create a `groovy-uno.jar` file
in build/libs.

## Using in your project

Once built (see above), add `groovy-uno.jar` as a dependency in your project. You'll also need to add
the following .jars from your Open/LibreOffice install (the directory after each file is the directory
that they are installed into on a typical debian install of LibreOffice):

 * `juh.jar`
 * `ridl.jar`
 * `jurt.jar`
 * `unoil.jar` (/usr/lib/libreoffice/program/classes/unoil.jar)

Here are these dependencies in a gradle friendly format:

```
dependencies {
	compile "org.openoffice:juh:3.2.1"
	compile "org.openoffice:ridl:3.2.1"
	compile "org.openoffice:unoil:3.2.1"
	compile "org.openoffice:jurt:3.2.1"
}
```

## Examples

### Connecting to running office

#### Java UNO

```java
XComponentContext context           = Bootstrap.bootstrap();
XMultiServiceFactory serviceFactory = context.getServiceManager();
XComponentContext context           = uno.util.Bootstrap.bootstrap();
XComponentLoader loader             = (XComponentLoader) UnoRuntime.queryInterface( XComponentLoader.class, serviceFactory.createInstanceWithContext( "com.sun.star.frame.Desktop", context ) );
XComponent component                = loader.loadComponentFromURL( "private:factory/scalc", "_blank", 0, new PropertyValue[0] );
XSpreadsheetDocument document       = (XSpreadsheetDocument) UnoRuntime.queryInterface( XSpreadsheetDocument.class, component );
```

#### groovy-uno

```groovy
SpreadsheetConnector connector = new SpreadsheetConnector()
XSpreadsheetDocument document  = connector.open()
```

### Get spreadsheet at index

#### Java UNO

```java
XSpreadsheets sheets     = document.getSheets();
XIndexAccess indexAccess = (XIndexAccess) UnoRuntime.queryInterface( XIndexAccess.class, sheets );
XSpreadsheet sheet       = (XSpreadsheet) UnoRuntime.queryInterface( XSpreadsheet.class, indexAccess.getByIndex( index ) );
```

#### groovy-uno

```groovy
XSpreadsheet sheet = document[ index ]
```

### Setting value at position

#### Java UNO

```java
XSpreadsheet sheet = ...;
sheet.getCellRangeByName( cellName ).getCellByPosition( 0, 0 ).setValue( value );
```

#### groovy-uno

```groovy
XSpreadsheet sheet = ...
sheet[ cellName ] << value
or
sheet[ cellName ].data = value
```

## Gotchas

UNO spreadsheet cells can be assigned by setting either their "data" or their "formula". If you want to populate a cell
with a string, you will need to assign it to the "formula", as "data" can only be numeric. This means that instead of
`sheet[ "A1" ] << 13.2` (which will work fine because 13.2 is being assigned to the "data" of the cell, you will need 
to use `sheet[ "A1" ].formula = "Random text"`.
