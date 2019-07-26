package com.bmp.tests.oo;

import com.sun.star.beans.PropertyValue;
import com.sun.star.bridge.UnoUrlResolver;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.io.XInputStream;
import com.sun.star.io.XOutputStream;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lib.uno.adapter.ByteArrayToXInputStreamAdapter;
import com.sun.star.lib.uno.adapter.OutputStreamToXOutputStreamAdapter;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

public class TestOpenOffice {

    private static final String BASE_PATH = "/Users/brunomarques/IdeaProjects/tests/ootests/src/main/resources/";

    public static void main(String[] args) throws Exception {
        XComponentContext xComponentContext = Bootstrap.createInitialComponentContext(null);
        XUnoUrlResolver xUnoUrlResolver = UnoUrlResolver.create(xComponentContext);

        Object initialObject = xUnoUrlResolver.resolve("uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager");
        XMultiComponentFactory xMultiComponentFactory = (XMultiComponentFactory) UnoRuntime.queryInterface(XMultiComponentFactory.class, initialObject);

        Object oDesktop = xMultiComponentFactory.createInstanceWithContext("com.sun.star.frame.Desktop", xComponentContext);
        XComponentLoader xComponentLoader = UnoRuntime.queryInterface(com.sun.star.frame.XComponentLoader.class, oDesktop);

        ///
        byte[] array = Files.readAllBytes(Paths.get(BASE_PATH + "Sample medium docx.docx"));
        XInputStream xInputStream = new ByteArrayToXInputStreamAdapter(array);

        PropertyValue[] conversionProperties = new PropertyValue[2];
        conversionProperties[0] = new PropertyValue();
        conversionProperties[1] = new PropertyValue();

        conversionProperties[0].Name = "InputStream";
        conversionProperties[0].Value = xInputStream;
        conversionProperties[1].Name = "Hidden";
        conversionProperties[1].Value = new Boolean(true);

        XComponent xComponent = xComponentLoader.loadComponentFromURL("private:stream", "_blank", 0, conversionProperties);

        ////
        XOutputStream outputStream = new OutputStreamToXOutputStreamAdapter(new FileOutputStream(BASE_PATH + "test.pdf"));

        conversionProperties[0].Name = "OutputStream";
        conversionProperties[0].Value = outputStream;
        conversionProperties[1].Name = "FilterName";
        conversionProperties[1].Value = "writer_pdf_Export";

        XStorable xStorable = (XStorable) UnoRuntime.queryInterface(XStorable.class, xComponent);
        xStorable.storeToURL("private:stream", conversionProperties);
    }
}
