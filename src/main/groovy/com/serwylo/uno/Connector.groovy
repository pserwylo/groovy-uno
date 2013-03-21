package com.serwylo.uno

import ooo.connector.BootstrapConnector
import com.sun.star.frame.XComponentLoader
import com.sun.star.lang.XMultiComponentFactory
import com.sun.star.uno.UnoRuntime
import com.sun.star.uno.XComponentContext
import ooo.connector.BootstrapSocketConnector
import ooo.connector.server.OOoServer
import ooo.connector.server.OOoServerPath

class Connector {

	private static boolean hasConnected = false

	private static BootstrapSocketConnector bootstrapSocketConnector
	private static OOoServer server

	private static XComponentContext context
    private static XMultiComponentFactory componentFactory
	private static XComponentLoader loader

	private static String NEW_PREFIX = "private:factory/"

	private static boolean connect() {
		if ( !hasConnected ) {
			List<String> options = OOoServer.getDefaultOOoOptions() + [ "--nofirststartwizard" ];
			server                   = new OOoServer(new OOoServerPath("/usr/bin/soffice"), options);
			bootstrapSocketConnector = new BootstrapSocketConnector(server);
			context                  = bootstrapSocketConnector.connect();
			componentFactory         = context.serviceManager
			hasConnected             = true
        }
		return hasConnected
	}

	private static void disconnect() {
		bootstrapSocketConnector.disconnect();
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
