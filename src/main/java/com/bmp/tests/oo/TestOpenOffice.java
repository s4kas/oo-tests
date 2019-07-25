package com.bmp.tests.oo;

import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.io.XInputStream;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lib.uno.adapter.ByteArrayToXInputStreamAdapter;
import com.sun.star.sdbc.XCloseable;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;

public class TestOpenOffice {

    public static void main(String[] args) throws BootstrapException, Exception, IOException {
        XComponentContext xComponentContext = Bootstrap.bootstrap();
        XMultiComponentFactory xMultiComponentFactory = xComponentContext.getServiceManager();

        Object oDesktop = xMultiComponentFactory.createInstanceWithContext("com.sun.star.frame.Desktop", xComponentContext);
        XComponentLoader xComponentLoader = UnoRuntime.queryInterface(com.sun.star.frame.XComponentLoader.class, oDesktop);

        ///
        byte[] array = Files.readAllBytes(Paths.get("classpath:Sample medium docx.docx"));
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

        OutputStream outputStream = new FileOutputStream("classpath:/test.pdf");

        conversionProperties[0].Name = "OutputStream";
        conversionProperties[0].Value = outputStream;
        conversionProperties[1].Name = "FilterName";
        conversionProperties[1].Value = "writer_pdf_Export";

        XStorable xStorable = (XStorable) UnoRuntime.queryInterface(XStorable.class,xComponent);
        xStorable.storeToURL("private:stream", conversionProperties);

        XCloseable xCloseable = (XCloseable) UnoRuntime.queryInterface(XCloseable.class,xComponent);
        xCloseable.close();
    }
}
