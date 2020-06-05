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
 * capitalise.groovy
 * Change the case of a selection, or current word from upper case,
 * to first char upper case, to all lower case to upper case...
 * derived from capitalise.bsh
 */

import com.sun.star.frame.XModel
import com.sun.star.frame.XController
import com.sun.star.view.XSelectionSupplier
import com.sun.star.container.XIndexAccess
import com.sun.star.text.XText
import com.sun.star.text.XTextRange
import com.sun.star.text.XWordCursor
import com.sun.star.script.provider.XScriptContext
import org.openoffice.guno.UnoExtension // the Groovy UNO Extension

// return the new string based on the string passed in
String getNewString(String theString) {

    String newString
    if(theString == null || theString.isEmpty()) {
        return newString;
    }
    // should we tokenize on "."?
    if (Character.isUpperCase(theString.charAt(0)) && theString.length() >= 2 && Character.isUpperCase(theString.charAt(1))) {
        // first two chars are UC => first UC, rest LC
        newString = theString.substring(0, 1).toUpperCase() + theString.substring(1).toLowerCase()
    } else if (Character.isUpperCase(theString.charAt(0))) { // first char UC => all to LC
        newString = theString.toLowerCase()
    } else { // all to UC.
        newString = theString.toUpperCase()
    }
    return newString
}

//the method that does the work
void capitalise(XIndexAccess xIndexAccess, XSelectionSupplier xSelectionSupplier) {

    // get the number of regions selected
    int count = xIndexAccess.getCount()
    if (count >= 1) { //ie we have a selection
        for (i = 0; i < count; i++) {
            // get the i-th region selected
            XTextRange xTextRange = xIndexAccess.getByIndex(i).guno(XTextRange.class)
            println("string: ${xTextRange.getString()}")

            // get the selected string
            theString = xTextRange.getString()
            if (theString.length() == 0) {
                // sadly we can have a selection where nothing is selected
                // in this case we get the XWordCursor and make a selection!
                XText xText = xTextRange.getText()
                XWordCursor xWordCursor = xText.createTextCursorByRange(xTextRange).guno(XWordCursor.class)
                // move the Word cursor to the start of the word if its not
                // already there
                if (!xWordCursor.isStartOfWord()) {
                    xWordCursor.gotoStartOfWord(false)
                }
                // move the cursor to the next word, selecting all chars
                // in between
                xWordCursor.gotoNextWord(true)
                // get the selected string
                theString = xWordCursor.getString()
                // get the new string
                newString = getNewString(theString)
                if (newString != null) {
                    // set the new string
                    xWordCursor.setString(newString)
                    // keep the current selection
                    xSelectionSupplier.select(xWordCursor)
                }
            } else {
                newString = getNewString(theString)
                if (newString != null) {
                    // set the new string
                    xTextRange.setString(newString)
                    // keep the current selection
                    xSelectionSupplier.select(xTextRange)
                }
            }

        }
    }
}

// The XSCRIPTCONTEXT variable is of type XScriptContext and is available to
// all Groovy scripts executed by the Script Framework
XModel xModel = XSCRIPTCONTEXT.getDocument()
XController xController = xModel.getCurrentController()

//the writer controller impl supports the css.view.XSelectionSupplier interface
XSelectionSupplier xSelectionSupplier = xController.guno(XSelectionSupplier.class)

//see section 7.5.1 of developers' guide
// XIndexAccess xIndexAccess = (XIndexAccess) UnoRuntime.queryInterface(XIndexAccess.class, xSelectionSupplier.getSelection());
XIndexAccess xIndexAccess = xSelectionSupplier.getSelection().guno(XIndexAccess.class)

//call the method that does the work
capitalise(xIndexAccess, xSelectionSupplier)
return 0
