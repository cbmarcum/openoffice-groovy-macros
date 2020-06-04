/* ************************************************************
 *
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 *
 *********************************************************** */

/*
 * exportsheetstohtml.groovy
 * When this script is run on an existing, saved, spreadsheet,
 * eg. /home/testuser/myspreadsheet.sxc, the script will export
 * each sheet to a separate html file,
 * eg. /home/testuser/myspreadsheet_sheet1.html,
 * /home/testuser/myspreadsheet_sheet2.html etc
 * derived from exportsheetstohtml.js
 */

import com.sun.star.uno.UnoRuntime
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.container.XIndexAccess
import com.sun.star.beans.XPropertySet
import com.sun.star.beans.PropertyValue
import com.sun.star.util.XModifiable
import com.sun.star.frame.XStorable
import com.sun.star.frame.XModel
import com.sun.star.uno.AnyConverter
import com.sun.star.uno.Type
import java.lang.System
import org.openoffice.guno.UnoExtension

// the Groovy UNO Extension

//get the document object from the scripting context
XModel xModel = XSCRIPTCONTEXT.getDocument()

//get the XSpreadsheetDocument interface from the document
// xSDoc = UnoRuntime.queryInterface(XSpreadsheetDocument, xModel)
xSDoc = xModel.guno(XSpreadsheetDocument.class)

//get the XModel interface from the document
// xModel = UnoRuntime.queryInterface(XModel,xModel);

//get the XIndexAccess interface used to access each sheet 
// xSheetsIndexAccess = UnoRuntime.queryInterface(XIndexAccess, xSDoc.getSheets())
xSheetsIndexAccess = xSDoc.getSheets().guno(XIndexAccess.class)

//get the XStorable interface used to save the document
// xStorable = UnoRuntime.queryInterface(XStorable,xSDoc);
xStorable = xSDoc.guno(XStorable.class)

//get the XModifiable interface used to indicate if the document has been 
//changed
// xModifiable = UnoRuntime.queryInterface(XModifiable,xSDoc);
xModifiable = xSDoc.guno(XModifiable.class)

//set up an array of PropertyValue objects used to save each sheet in the 
//document
PropertyValue[] storeProps = new PropertyValue[1] //PropertyValue[1]
storeProps[0] = new PropertyValue()
storeProps[0].Name = "FilterName"
storeProps[0].Value = "HTML (StarCalc)"
storeUrl = xModel.getURL()
storeUrl = storeUrl.substring(0, storeUrl.lastIndexOf('.'))

//set only one sheet visible, and store to HTML doc
for (int i = 0; i < xSheetsIndexAccess.getCount(); i++ ) {
    setAllButOneHidden(xSheetsIndexAccess, i)
    xModifiable.setModified(false)
    xStorable.storeToURL(storeUrl + "_sheet" + (i + 1) + ".html", storeProps)
}

// now set all visible again
for (int i = 0; i < xSheetsIndexAccess.getCount(); i++ ) {
    XPropertySet xPropSet = AnyConverter.toObject(new Type(XPropertySet), xSheetsIndexAccess.getByIndex(i))
    xPropSet.setPropertyValue("IsVisible", true)
}

static void setAllButOneHidden(xSheetsIndexAccess, vis) {
    //System.err.println("count="+xSheetsIndexAccess.getCount())
    //get an XPropertySet interface for the vis-th sheet
    XPropertySet xPropSet = AnyConverter.toObject(new Type(XPropertySet), xSheetsIndexAccess.getByIndex(vis))
    //set the vis-th sheet to be visible
    xPropSet.setPropertyValue("IsVisible", true)
    // set all other sheets to be invisible
    for (int i = 0; i < xSheetsIndexAccess.getCount(); i++ )
    {
        xPropSet = AnyConverter.toObject(new Type(XPropertySet), xSheetsIndexAccess.getByIndex(i))
        if (i != vis) {
            xPropSet.setPropertyValue("IsVisible", false)
        }
    }
} 
