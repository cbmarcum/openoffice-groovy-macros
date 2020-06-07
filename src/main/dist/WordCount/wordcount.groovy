/* *************************************************************
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
 * wordcount.groovy
 * Provides a word count of the selected text in A Writer document.
 * derived from wordcount.bsh
 */
import java.awt.BorderLayout
import java.awt.event.ActionEvent
import java.awt.event.ActionListener
import javax.swing.JButton
import javax.swing.JFrame
import javax.swing.JLabel

// import com.sun.star.uno.UnoRuntime;
import com.sun.star.frame.XModel
import com.sun.star.view.XSelectionSupplier
import com.sun.star.container.XIndexAccess
import com.sun.star.text.XText
import com.sun.star.text.XTextRange
import com.sun.star.script.provider.XScriptContext

// the Groovy UNO Extension
import org.openoffice.guno.UnoExtension


// display the count in a Swing dialog
void doDisplay(Integer numWords) {
    JLabel wordsLabel = new JLabel("Word count = " + numWords)
    JButton closeButton = new JButton("Close")
    JFrame frame = new JFrame("Word Count")
    closeButton.addActionListener(new ActionListener() {
        void actionPerformed(ActionEvent e) {
            frame.setVisible(false)
        }
    })
    frame.getContentPane().setLayout(new BorderLayout())
    frame.getContentPane().add(wordsLabel, BorderLayout.CENTER)
    frame.getContentPane().add(closeButton, BorderLayout.SOUTH)
    frame.pack()
    frame.setSize(190, 90)
    frame.setLocation(430, 430)
    frame.setVisible(true)
}

Integer wordcount(XIndexAccess xIndexAccess) {

    Integer result = 0

    // iterate through each of the selections
    Integer count = xIndexAccess.getCount()
    for (int i = 0; i < count; i++) {
        // get the XTextRange of the selection
        // XTextRange xTextRange = (XTextRange)UnoRuntime.queryInterface(XTextRange.class, xIndexAccess.getByIndex(i))
        XTextRange xTextRange = xIndexAccess.getByIndex(i).guno(XTextRange.class)

        // println("string: "+xTextRange.getString());
        // use the standard J2SE delimiters to tokenize the string
        // obtained from the XTextRange
        StringTokenizer strTok = new StringTokenizer(xTextRange.getString())
        result += strTok.countTokens();
    }

    doDisplay(result);
    return result;
}

// The XSCRIPTCONTEXT variable is of type XScriptContext and is available to
// all Groovy scripts executed by the Script Framework
XModel xModel = XSCRIPTCONTEXT.getDocument()

//the writer controller impl supports the css.view.XSelectionSupplier interface
XSelectionSupplier xSelectionSupplier = xModel.getCurrentController().guno(XSelectionSupplier.class)

//see section 7.5.1 of developers' guide
// the getSelection provides an XIndexAccess to the one or more selections
XIndexAccess xIndexAccess = xSelectionSupplier.getSelection().guno(XIndexAccess.class)

Integer count = wordcount(xIndexAccess)
println("count = " + count)

// Groovy OpenOffice scripts should always return 0
return 0
