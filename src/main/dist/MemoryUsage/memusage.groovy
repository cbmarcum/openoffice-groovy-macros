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
 * memusage.groovy
 * Creates a spreadsheet with the current memory usage
 * statistics for the Java Virtual Machine
 * derived from memusage.bsh
 */

import com.sun.star.uno.UnoRuntime
import com.sun.star.uno.AnyConverter
import com.sun.star.uno.Type
import com.sun.star.uno.XInterface
import com.sun.star.lang.XComponent
import com.sun.star.lang.XMultiServiceFactory
import com.sun.star.frame.XComponentLoader
import com.sun.star.frame.XDesktop
import com.sun.star.document.XEmbeddedObjectSupplier
import com.sun.star.awt.ActionEvent
import com.sun.star.awt.Rectangle
import com.sun.star.beans.XPropertySet
import com.sun.star.beans.PropertyValue

import com.sun.star.container.*
import com.sun.star.chart.*
import com.sun.star.table.*
import com.sun.star.sheet.*
import com.sun.star.script.provider.XScriptContext
// the Groovy UNO Extension
import org.openoffice.guno.UnoExtension


static void addData(XSpreadsheet sheet, String date, total, Long free) {
    // set the labels
    sheet.getCellByPosition(0, 0).setFormula("Used");
    sheet.getCellByPosition(0, 1).setFormula("Free");
    sheet.getCellByPosition(0, 2).setFormula("Total");

    // set the values in the cells
    sheet.getCellByPosition(1, 0).setValue(total - free);
    sheet.getCellByPosition(1, 1).setValue(free);
    sheet.getCellByPosition(1, 2).setValue(total);
}

static void addChart(XSpreadsheet sheet) {
    Rectangle rect = new Rectangle()
    rect.X = 500
    rect.Y = 3000;
    rect.Width = 10000
    rect.Height = 8000

    XCellRange range = (XCellRange) UnoRuntime.queryInterface(XCellRange.class, sheet)
    XCellRange myRange = range.getCellRangeByName("A1:B2")

    // XCellRangeAddressable rangeAddr = (XCellRangeAddressable)UnoRuntime.queryInterface(XCellRangeAddressable.class, myRange)
    XCellRangeAddressable rangeAddr = myRange.guno(XCellRangeAddressable.class)
    CellRangeAddress myAddr = rangeAddr.getRangeAddress()

    CellRangeAddress[] addr = new CellRangeAddress[1]
    addr[0] = myAddr

    // XTableChartsSupplier supp = (XTableChartsSupplier)UnoRuntime.queryInterface(XTableChartsSupplier.class, sheet)
    XTableChartsSupplier supp = sheet.guno(XTableChartsSupplier.class)
    XTableCharts charts = supp.getCharts()
    charts.addNewByName("Example", rect, addr, false, true)

    try {
        Thread.sleep(3000);
    } catch (java.lang.InterruptedException e) {
    }

    // get the diagram and Change some of the properties
    // XNameAccess chartsAccess = (XNameAccess)UnoRuntime.queryInterface(XNameAccess.class, charts)
    XNameAccess chartsAccess = charts.guno(XNameAccess.class)

    // XTableChart tchart = (XTableChart)UnoRuntime.queryInterface(XTableChart.class, chartsAccess.getByName("Example"))
    XTableChart tchart = chartsAccess.getByName("Example").guno(XTableChart.class)

    // XEmbeddedObjectSupplier eos = (XEmbeddedObjectSupplier)UnoRuntime.queryInterface(XEmbeddedObjectSupplier.class, tchart)

    XEmbeddedObjectSupplier eos = tchart.guno(XEmbeddedObjectSupplier.class)
    XComponent xifc = eos.getEmbeddedObject()

    // XChartDocument xChart = (XChartDocument)UnoRuntime.queryInterface(XChartDocument.class, xifc)
    XChartDocument xChart = xifc.guno(XChartDocument.class)

    // XMultiServiceFactory xDocMSF = (XMultiServiceFactory)UnoRuntime.queryInterface(XMultiServiceFactory.class, xChart);
    XMultiServiceFactory xDocMSF = xChart.guno(XMultiServiceFactory.class)

    XInterface diagObject = xDocMSF.createInstance("com.sun.star.chart.PieDiagram")

    // XDiagram xDiagram = (XDiagram)UnoRuntime.queryInterface(XDiagram.class, diagObject)
    XDiagram xDiagram = diagObject.guno(XDiagram.class)
    xChart.setDiagram(xDiagram)

    // XPropertySet propset = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xChart.getTitle())
    XPropertySet propset = xChart.getTitle().guno(XPropertySet.class)
    propset.setPropertyValue("String", "JVM Memory Usage")
}

Runtime runtime = Runtime.getRuntime()
Random generator = new Random()
Date date = new Date()

// allocate a random number of bytes so that the data changes
int len = (int) (generator.nextFloat() * runtime.freeMemory() / 5)
byte[] bytes = new byte[len]

XDesktop xDesktop = XSCRIPTCONTEXT.getDesktop()
XComponentLoader loader = xDesktop.guno(XComponentLoader.class)
XComponent comp = loader.loadComponentFromURL("private:factory/scalc", "_blank", 4, new PropertyValue[0])
XSpreadsheetDocument doc = comp.guno(XSpreadsheetDocument.class)
XSpreadsheet sheet = doc.getSheetByIndex(0) // Groovy UNO Extension method

addData(sheet, date.toString(), runtime.totalMemory(), runtime.freeMemory());
addChart(sheet)

return 0
