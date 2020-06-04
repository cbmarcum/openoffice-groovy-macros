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
 * helloworld.groovy
 * Adds "Hello World (in Groovy)" at cursor location in a text document.
 * derived from helloworld.bsh
 */

import com.sun.star.frame.XModel
import com.sun.star.text.XText
import com.sun.star.text.XTextDocument
import com.sun.star.text.XTextRange
import org.openoffice.guno.UnoExtension // the Groovy UNO Extension

// set the output text string
String output = "Hello World (in Groovy)"

// get the document model from the scripting context which is made available to all scripts
XModel xModel = XSCRIPTCONTEXT.getDocument()

// get the XTextDocument interface
XTextDocument xTextDoc = xModel.guno(XTextDocument.class)

//get the XText interface
XText xText = xTextDoc.getText()

// get an (empty) XTextRange at the end of the document
XTextRange xTextRange = xText.getEnd()

// sets the text of the text range
xTextRange.setString( output )

// Groovy OpenOffice scripts should always return 0
return 0