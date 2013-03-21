package com.serwylo.uno

import ooo.connector.BootstrapConnector
import com.sun.star.comp.helper.Bootstrap
import com.sun.star.frame.XComponentLoader
import com.sun.star.lang.XMultiComponentFactory
import com.sun.star.uno.UnoRuntime
import com.sun.star.uno.XComponentContext

class Connector {

	private static boolean hasConnected = false

	private static XComponentContext context
    private static XMultiComponentFactory componentFactory
	private static XComponentLoader loader

	private static String NEW_PREFIX = "private:factory/"

	private static boolean connect() {
		if ( !hasConnected ) {
			context = BootstrapConnector.bootstrap( "/usr/bin/", "", "" )
			componentFactory = context.serviceManager
			hasConnected = true
        }
		return hasConnected
	}

	protected Connector() {
		connect()
	}

	public XComponentContext getContext() {
		context
	}

	public XMultiComponentFactory getComponentFactory() {
		componentFactory
	}

	protected String generateNewUrl( String component ) {
		"$NEW_PREFIX$component"
	}

	protected Boolean isNewUrl( String path ) {
		path.startsWith( NEW_PREFIX )
	}

	protected String sanitizePath( String path ) {
		if ( !isNewUrl( path ) ) {
			if ( !path.startsWith( "file://" ) ) {
				path = new File( path ).absolutePath
				path = "file://$path"
			}
		}
		return path
	}

	protected XComponentLoader getLoader() {
		if ( loader == null ) {
			loader = UnoRuntime.queryInterface(
				XComponentLoader.class,
				componentFactory.createInstanceWithContext("com.sun.star.frame.Desktop", context))
		}
		return loader
	}

}
