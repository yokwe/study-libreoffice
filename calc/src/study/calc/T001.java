package study.calc;

import java.io.File;

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
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;

// Open blank writer window
//   https://technology.amis.nl/2006/06/22/getting-started-with-the-openofficeorg-api-part-i-some-basic-writer-operations/

// http://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1Spreadsheet.html
// https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Spreadsheet_Services_-_Overview

public class T001 {

	public static void main(String[] args) {
		// Get the remote office component context
		File targetFile = new File("/home/hasegawa/Dropbox/Trade/投資損益計算_2016_SAVE.ods");
		XComponent component = null;
		System.out.println("START");
		try {
			{
				String[] bootstrapOptions = {
						"--minimized",
						"--headless",
						"--invisible",
				};
				PropertyValue[] props = new PropertyValue[] {
						// Set document as read only
						new PropertyValue("ReadOnly", 0, true, PropertyState.DIRECT_VALUE),
						// Update linked reference
						new PropertyValue("UpdateDocMode", 0, UpdateDocMode.QUIET_UPDATE, PropertyState.DIRECT_VALUE),
				};
				String targetURL = targetFile.toURI().toURL().toExternalForm();
				System.out.format("targetURL = %s\n", targetURL);

				XComponentContext componentContext = Bootstrap.bootstrap(bootstrapOptions);
				XMultiComponentFactory serviceManager = componentContext.getServiceManager();
				Object desktop = serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop", componentContext);
				XComponentLoader componentLoader = UnoRuntime.queryInterface(XComponentLoader.class, desktop);
				component = componentLoader.loadComponentFromURL(targetURL, "_blank", 0, props);
			}
			XSpreadsheets sheets = UnoRuntime.queryInterface(XSpreadsheetDocument.class, component).getSheets();

			// Enumerate all sheet using index access
			{
				XIndexAccess indexAccess = UnoRuntime.queryInterface(XIndexAccess.class, sheets);
				System.out.println("index = " + indexAccess.getCount());

				for (int i = 0; i < indexAccess.getCount(); i++) {
					XSpreadsheet sheet = UnoRuntime.queryInterface(XSpreadsheet.class, indexAccess.getByIndex(i));
					XNamed named = UnoRuntime.queryInterface(XNamed.class, sheet);
					System.out.println("Index " + i + " - " + named.getName());
				}
			}

			// Enumerate all sheet using name access
			{
				XNameAccess nameAccess = UnoRuntime.queryInterface(XNameAccess.class, sheets);
				System.out.println("name = " + nameAccess.getElementNames().length);
				for (String name : nameAccess.getElementNames()) {
					XSpreadsheet sheet = UnoRuntime.queryInterface(XSpreadsheet.class, nameAccess.getByName(name));
					XNamed named = UnoRuntime.queryInterface(XNamed.class, sheet);
					System.out.println("Name " + name + " - " + named.getName());
				}
			}

			// Enumerate all sheet using enumeration
			{
				XEnumerationAccess enumerateAccess = UnoRuntime.queryInterface(XEnumerationAccess.class, sheets);
				XEnumeration eumeration = enumerateAccess.createEnumeration();
				while (eumeration.hasMoreElements()) {
					XSpreadsheet sheet = UnoRuntime.queryInterface(XSpreadsheet.class, eumeration.nextElement());
					XNamed named = UnoRuntime.queryInterface(XNamed.class, sheet);
					System.out.println("Enum " + named.getName());
				}
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (component != null) {
				XCloseable closeable = null;
				{
					XModel xModel = UnoRuntime.queryInterface(XModel.class, component);
					if (xModel != null) {
						closeable = UnoRuntime.queryInterface(XCloseable.class, xModel);
					}
				}
				
				if (closeable != null) {
					try {
						System.out.println("Invoke close");
						closeable.close(true);
					} catch (CloseVetoException e) {
						System.out.println("CloseVetoException " + e.toString());
					}
				} else {
					System.out.println("Invoke disponse");
					component.dispose();
				}
			}
			System.out.println("STOP");
			
			// Explicitly exit process
			System.exit(0);
		}
	}

}
