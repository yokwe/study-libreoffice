package study.calc;

import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


// Open blank writer window
//   https://technology.amis.nl/2006/06/22/getting-started-with-the-openofficeorg-api-part-i-some-basic-writer-operations/

public class T001 {

	public static void main(String[] args) {
		// Get the remote office component context
		System.out.println("START");
		try {
			XComponentContext xLocalContext = Bootstrap.bootstrap();
			XMultiComponentFactory xLocalServiceManager = xLocalContext.getServiceManager();
			Object desktop = xLocalServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", xLocalContext);
			XComponentLoader xComponentLoader = (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class, desktop);
			PropertyValue[] pPropValues = new PropertyValue[0];
			// Use "private:factory/scalc" for calc
			XComponent xComponent = xComponentLoader.loadComponentFromURL("private:factory/swriter", "_blank", 0, pPropValues);
		} catch (Throwable e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		System.out.println("STOP");
	}

}
