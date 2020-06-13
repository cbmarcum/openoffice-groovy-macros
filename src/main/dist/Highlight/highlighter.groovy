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
 * highlighter.groovy
 * A Swing Dialog with button press events to highlight text.
 * derived from highlighter.bsh
 *
 */

import com.sun.star.beans.PropertyValue
import com.sun.star.script.provider.XScriptContext
import com.sun.star.text.XTextDocument
import com.sun.star.util.*
import com.sun.star.util.XPropertyReplace
import com.sun.star.util.XReplaceDescriptor
import com.sun.star.util.XReplaceable
import groovy.transform.Field

import javax.swing.*
import java.awt.*
import java.awt.event.ActionEvent
import java.awt.event.ActionListener

// the Groovy UNO Extension
import org.openoffice.guno.UnoExtension

@Field String searchKey = ""
@Field XTextDocument xTextDocument = XSCRIPTCONTEXT.getDocument().guno(XTextDocument.class)
@Field JTextField findTextBox = new JTextField(20)
@Field JFrame frame = new JFrame("Highlight Text")

Integer replaceText(String searchKey, Integer color, Boolean bold) {

    Integer result = 0

    try {
        // Create an XReplaceable object and an XReplaceDescriptor
        XReplaceable replaceable = xTextDocument.guno(XReplaceable.class)

        XReplaceDescriptor descriptor = replaceable.createReplaceDescriptor()

        // Gets a XPropertyReplace object for altering the properties
        // of the replaced text
        XPropertyReplace xPropertyReplace = descriptor.guno(XPropertyReplace.class)

        // Sets the replaced text property fontweight value to Bold or Normal 
        PropertyValue wv = null
        if (bold) {
            wv = new PropertyValue("CharWeight", -1,
                new Float(com.sun.star.awt.FontWeight.BOLD),
                com.sun.star.beans.PropertyState.DIRECT_VALUE)
        }
        else {
            wv = new PropertyValue("CharWeight", -1,
                new Float(com.sun.star.awt.FontWeight.NORMAL),
                com.sun.star.beans.PropertyState.DIRECT_VALUE)
        }

        // Sets the replaced text property color value to RGB color parameter
        PropertyValue cv = new PropertyValue("CharColor", -1, new Integer(color),
            com.sun.star.beans.PropertyState.DIRECT_VALUE)

        // Apply the properties
        PropertyValue[] props = [cv, wv]
        xPropertyReplace.setReplaceAttributes(props)

        // Only matches whole words and case sensitive
        descriptor.setPropertyValue("SearchCaseSensitive", new Boolean(true))
        descriptor.setPropertyValue("SearchWords", new Boolean(true))

        // Replaces all instances of searchKey with new Text properties
        // and gets the number of instances of the searchKey 
        descriptor.setSearchString(searchKey)
        descriptor.setReplaceString(searchKey)
        result = replaceable.replaceAll(descriptor)

    }
    catch (Exception e) {
        println(e.toString())
    }

    return result
}



// Create a JButton and add an ActionListener
// When clicked the value for the searchKey is read and passed to replaceText
ActionListener myListener = new ActionListener() {
    void actionPerformed(ActionEvent e) {
        searchKey = findTextBox.getText()

        if(searchKey.isEmpty()) {
            JOptionPane.showMessageDialog(null,
                "No text entered for search",
                "No text", JOptionPane.INFORMATION_MESSAGE)
        }
        else {
            // highlight the text in red
            Color cRed = new Color(255, 0, 0)
            Integer red = cRed.getRGB()
            Integer num = replaceText(searchKey, red, true)

            if(num > 0) {
                Integer response = JOptionPane.showConfirmDialog(null,
                    "${searchKey} was found ${num} times\nDo you wish to keep the text highlighted?",
                    "Confirm highlight", JOptionPane.YES_NO_OPTION,
                    JOptionPane.QUESTION_MESSAGE)

                if (response == 1) {
                    Color cBlack = new Color(255, 255, 255)
                    Integer black = cBlack.getRGB()
                    replaceText(searchKey, black, false)
                }
            } else {
                JOptionPane.showMessageDialog(null,
                    "No matches were found", "Not found",
                     JOptionPane.INFORMATION_MESSAGE)
            }
        }
    }
}

ActionListener exitListener = new ActionListener() {
    void actionPerformed(ActionEvent e) {
        frame.dispose()
    }
}

JButton searchButton = new JButton("Highlight")
searchButton.addActionListener(myListener)

JButton exitButton = new JButton("Exit")
exitButton.addActionListener(exitListener)

JPanel buttonPanel = new JPanel()
buttonPanel.setLayout(new FlowLayout())
buttonPanel.add(searchButton)
buttonPanel.add(exitButton)

// create a JPanel containing one JTextField for the search text
JPanel searchPanel = new JPanel()
searchPanel.setLayout(new FlowLayout())

JLabel findWhat = new JLabel("Find What: ")
searchPanel.add(findWhat)
searchPanel.add(findTextBox)

// dd a window listener to the frame
frame.setSize(350,130)
frame.setLocation(430,430)
frame.setResizable(false)
// add the panel and button to the frame
frame.getContentPane().setLayout(new GridLayout(2,1,10,10))
frame.getContentPane().add(searchPanel)
frame.getContentPane().add(buttonPanel)
frame.setVisible(true)
frame.pack()
