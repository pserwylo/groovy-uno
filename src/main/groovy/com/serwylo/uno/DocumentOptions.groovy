package com.serwylo.uno

import com.sun.star.beans.PropertyValue

class DocumentOptions {

	boolean hideFrame = true

	PropertyValue[] getProperties() {
		List<PropertyValue> documentProps = [ getPropertyHideFrame() ]
		documentProps.addAll( additionalProperties )
		(PropertyValue[])documentProps.toArray()
	}

	/**
	 * If this was an abstract class, this method would be abstract. But it turns out that sometimes you may just
	 * not care about the options for a specific document type.
	 * @return
	 */
	protected List<PropertyValue> getAdditionalProperties() { [] }

	private PropertyValue getPropertyHideFrame() {
		new PropertyValue( Name: "Hidden", Value: hideFrame )
	}

}
