package com.serwylo.uno

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

	private static boolean connect() {
		if ( !hasConnected ) {
            try {
                context = Bootstrap.bootstrap()
                componentFactory = context.serviceManager
				hasConnected = true
            } catch( Exception e ) {}
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
		return "private:factory/$component"
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
