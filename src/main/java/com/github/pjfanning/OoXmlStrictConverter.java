package com.github.pjfanning;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;
import javax.xml.stream.XMLStreamException;

public class OoXmlStrictConverter {

    public static void main(String[] args) {
        try {
            if (args.length > 0) {
                transform(args[0], "transformed.xlsx");
            } else {
                transform("SimpleStrict.xlsx", "Simple.xlsx");
                transform("SampleSS.strict.xlsx", "SampleSS.trans.xlsx");
                transform("sample.strict.xlsx", "sample.trans.xlsx");
                transform("SimpleNormal.xlsx", "SimpleNormal.transformed.xlsx");

                transformUsingStream("SimpleStrict.xlsx", "Simple.trans-stream.xlsx");
                transformUsingStream("SampleSS.strict.xlsx", "SampleSS.trans-stream.xlsx");
                transformUsingStream("sample.strict.xlsx", "sample.trans-stream.xlsx");
                transformUsingStream("SimpleNormal.xlsx", "SimpleNormal.trans-stream.xlsx");
            }
        } catch(Throwable t) {
            t.printStackTrace();
        }
    }

    private static void transformUsingStream(String inFile, String outFile)
            throws IOException {
        System.out.println("transforming using stream " + inFile + " to " + outFile);
        try (InputStream in = new FileInputStream(inFile);
            InputStream translator = new OoXmlStrictConverterInputStream(in);
            OutputStream out = new FileOutputStream(outFile)) {

            copy(translator, out);
        }
    }

    private static void transform(String inFile, String outFile) throws Exception {
        System.out.println("transforming " + inFile + " to " + outFile);
        transform(new FileInputStream(inFile), new FileOutputStream(outFile));
    }

    private static void transform(InputStream inFile, OutputStream outFile) throws Exception {
        try(ZipInputStream zis = new ZipInputStream(inFile);
                ZipOutputStream zos = new ZipOutputStream(outFile);) {

            boolean successfullyRead;
            do {
                successfullyRead = convertEntry(zis, zos);

            } while (successfullyRead);
        }
    }

    private static boolean convertEntry(ZipInputStream zis, ZipOutputStream zos)
            throws IOException {

        ZipEntry ze = zis.getNextEntry();
        if (ze == null) {
            return false;
        }

        ZipEntry newZipEntry = new ZipEntry(ze.getName());
        zos.putNextEntry(newZipEntry);

        if(OoXmlStrictConverterUtils.isXml(ze.getName())) {
            try {
                convertXml(OoXmlStrictConverterUtils.disableClose(zis), OoXmlStrictConverterUtils.disableClose(zos));
            } catch(Throwable t) {
                throw new IOException("Problem parsing " + ze.getName(), t);
            }
        } else {
            copy(zis, zos);
        }
        zis.closeEntry();
        zos.closeEntry();

        return true;
    }

    private static void convertXml(InputStream is, OutputStream os)
            throws XMLStreamException {

        try (XmlConverter xmlConverter = new XmlConverter(is, os)) {
            while (xmlConverter.convertNextElement()) {}
        }
    }

    private static void copy(InputStream inp, OutputStream out) throws IOException {
        byte[] buff = new byte[4096];
        int count;
        while ((count = inp.read(buff)) != -1) {
            if (count > 0) {
                out.write(buff, 0, count);
            }
        }
    }
}
