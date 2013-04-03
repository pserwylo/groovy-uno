package com.serwylo.uno

import com.sun.star.beans.PropertyValue

class DocumentOptions {

	boolean hideFrame = true

	boolean overwrite = false

	PropertyValue[] getProperties() {
		List<PropertyValue> documentProps = [
			getPropertyHideFrame(),
			// getPropertyOverwrite(),
		]
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

	private PropertyValue getPropertyOverwrite() {
		new PropertyValue( Name: "Overwrite", Value: overwrite )
	}

}
