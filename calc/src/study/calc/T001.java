package study.calc;

import com.sun.star.beans.PropertyState;
import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNamed;
import com.sun.star.document.UpdateDocMode;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


// Open blank writer window
//   https://technology.amis.nl/2006/06/22/getting-started-with-the-openofficeorg-api-part-i-some-basic-writer-operations/

// http://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1Spreadsheet.html
// https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Spreadsheet_Services_-_Overview

public class T001 {

	public static void main(String[] args) {
		// Get the remote office component context
		XComponent xComponent = null;
		System.out.println("START");
		try {
			XComponentContext xLocalContext = Bootstrap.bootstrap();
			XMultiComponentFactory xLocalServiceManager = xLocalContext.getServiceManager();
			Object desktop = xLocalServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", xLocalContext);
			XComponentLoader xComponentLoader = (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class, desktop);
			
            PropertyValue [] props = new com.sun.star.beans.PropertyValue[] {
                new PropertyValue("ReadOnly", 0, true, PropertyState.DIRECT_VALUE),
                new PropertyValue("UpdateDocMode", 0, UpdateDocMode.QUIET_UPDATE, PropertyState.DIRECT_VALUE),
            };

			String url = "file:///home/hasegawa/Dropbox/Trade/投資損益計算_2016_SAVE.ods";  // existing file
			xComponent = xComponentLoader.loadComponentFromURL(url, "_blank", 0, props);
			XSpreadsheets sheets = UnoRuntime.queryInterface(XSpreadsheetDocument.class, xComponent).getSheets();
			
			// Enumerate all sheet using index access
			{
				XIndexAccess indexAccess = UnoRuntime.queryInterface(XIndexAccess.class, sheets);
				System.out.println("index = " + indexAccess.getCount());
				
				for(int i = 0; i < indexAccess.getCount(); i++) {
					XSpreadsheet sheet = UnoRuntime.queryInterface(XSpreadsheet.class, indexAccess.getByIndex(i));
					XNamed named = UnoRuntime.queryInterface(XNamed.class, sheet);
					System.out.println("Index " + i + " - " + named.getName());
				}
			}
			
			// Enumerate all sheet using name access
			{
				XNameAccess nameAccess = UnoRuntime.queryInterface(XNameAccess.class, sheets);
				System.out.println("name = " + nameAccess.getElementNames().length);
				for(String name: nameAccess.getElementNames()) {
					XSpreadsheet sheet = UnoRuntime.queryInterface(XSpreadsheet.class, nameAccess.getByName(name));
					XNamed named = UnoRuntime.queryInterface(XNamed.class, sheet);
					System.out.println("Name " + name + " - " + named.getName());
				}
			}

			// Enumerate all sheet using enumeration
			{
				XEnumerationAccess enumerateAccess = UnoRuntime.queryInterface(XEnumerationAccess.class, sheets);
				XEnumeration eumeration = enumerateAccess.createEnumeration();
				while(eumeration.hasMoreElements()) {
					XSpreadsheet sheet = UnoRuntime.queryInterface(XSpreadsheet.class, eumeration.nextElement());
					XNamed named = UnoRuntime.queryInterface(XNamed.class, sheet);
					System.out.println("Enum " + named.getName());
				}
			}
		} catch (Throwable e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (xComponent != null) xComponent.dispose();
		}

		System.out.println("STOP");
	}

}
