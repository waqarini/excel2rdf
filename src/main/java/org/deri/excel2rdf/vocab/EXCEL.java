
package org.deri.excel2rdf.vocab;

import com.hp.hpl.jena.rdf.model.Model;
import com.hp.hpl.jena.rdf.model.ModelFactory;
import com.hp.hpl.jena.rdf.model.Property;
import com.hp.hpl.jena.rdf.model.Resource;

public class EXCEL {

    public static String getURI() {
        return uri;
    }
    protected static final String uri = "http://office.microsoft.com/excel#";
    private static Model m = ModelFactory.createDefaultModel();
    public static final Resource cell = m.createResource(uri + "cell");
    public static final Resource Sheet = m.createResource(uri + "xlsheet");
    public static final Property sheetname = m.createProperty(uri + "sheetname");
    public static final Property row = m.createProperty(uri + "row");
    public static final Property column = m.createProperty(uri + "column");
    public static final Property columnIndex = m.createProperty(uri + "columnIndex");
    public static final Property sheet = m.createProperty(uri + "sheet");
    public static final Property value = m.createProperty(uri + "value");
    public static final Property backgroundColor = m.createProperty(uri + "backgroundColor");
    public static final Property foregroundColor = m.createProperty(uri + "foregroundColor");
    public static final Property up = m.createProperty(uri + "up");
    public static final Property down = m.createProperty(uri + "down");
    public static final Property left = m.createProperty(uri + "left");
    public static final Property right = m.createProperty(uri + "right");
    public static final Property type = m.createProperty(uri + "type");
    public static final Property cellType = m.createProperty(uri + "cellType");
    public static final Property findNextUp = m.createProperty(uri + "findNextUp");
    public static final Property findNextDown = m.createProperty(uri + "findNextDown");
    public static final Property findNextLeft = m.createProperty(uri + "findNextLeft");
    public static final Property findNextRight = m.createProperty(uri + "findNextRight");
}
