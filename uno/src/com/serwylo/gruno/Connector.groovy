package com.serwylo.gruno

import com.sun.star.comp.helper.Bootstrap
import com.sun.star.frame.XComponentLoader
import com.sun.star.lang.XMultiComponentFactory
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.uno.UnoRuntime
import com.sun.star.uno.XComponentContext

class Connector {

	private boolean hasConnected = false

	private XComponentContext remoteContext
    private XMultiComponentFactory  remoteServiceManager
	private XComponentLoader loader

	public boolean connect() {
		if ( !hasConnected ) {
            try {
                remoteContext = Bootstrap.bootstrap()
                remoteServiceManager = remoteContext.serviceManager
				hasConnected = true
            } catch( Exception e ) {}
        }
		return hasConnected
	}

	public XComponentContext getContext() {
		remoteContext
	}

	public XMultiComponentFactory getServiceManager() {
		remoteServiceManager
	}

	protected String generateNewUrl( String component ) {
		return "private:factory/$component"
	}

	protected XComponentLoader getLoader() {
		if ( loader == null ) {
			loader = UnoRuntime.queryInterface(
				XComponentLoader.class,
				remoteServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", remoteContext))
		}
		loader
	}

}
