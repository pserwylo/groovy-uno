package com.serwylo.gruno.utils

import com.serwylo.gruno.Connector
import com.sun.star.beans.XIntrospection
import com.sun.star.uno.UnoRuntime

class Introspector {

	private Connector connector

	public Introspector( Connector connector ) {
		this.connector = connector
	}

	public void introspect( Object object ) {
		try {
			Object oMRI = connector.componentFactory.createInstanceWithContext( "mytools.Mri", connector.context );
			XIntrospection xIntrospection = UnoRuntime.queryInterface( XIntrospection.class, oMRI);

			xIntrospection.inspect( object );
		} catch ( com.sun.star.uno.Exception e ) {
			System.err.println();
		}
	}
}
