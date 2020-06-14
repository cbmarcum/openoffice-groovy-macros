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
 * messagebox.groovy
 * example of UNO message boxes using Groovy UNO methods.
 * display an information box with a default title and then
 * display a warning box with a defined title and ok button
 * set as default.
 */

import com.sun.star.awt.MessageBoxType
import com.sun.star.awt.XMessageBox
import com.sun.star.awt.MessageBoxButtons
import com.sun.star.uno.XComponentContext

// the Groovy UNO Extension
import org.openoffice.guno.UnoExtension

XComponentContext xContext = XSCRIPTCONTEXT.getComponentContext()

// infobox type ignores buttons and uses BUTTONS_OK
String infoMsg = "This in an informative message..."
Integer infoButtons = MessageBoxButtons.BUTTONS_OK
XMessageBox infoBox = xContext.getMessageBox(MessageBoxType.INFOBOX, infoButtons, infoMsg)
short infoBoxResult = infoBox.execute()

String warnMsg = "This is a warning message...\nYou should be careful."
Integer warnButtons = MessageBoxButtons.BUTTONS_OK_CANCEL + MessageBoxButtons.DEFAULT_BUTTON_OK
XMessageBox warningBox = xContext.getMessageBox(MessageBoxType.WARNINGBOX, warnButtons, warnMsg, "Warning Title")
short warnBoxResult = warningBox.execute()

// Groovy OpenOffice scripts should always return 0
return 0
