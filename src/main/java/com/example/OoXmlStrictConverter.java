package com.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilterInputStream;
import java.io.FilterOutputStream;
import java.io.IOException;
import java.util.Properties;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventFactory;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLEventWriter;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

public class OoXmlStrictConverter {

    private static final String INPUT_FILENAME = "SimpleStrict.xlsx";
    private static final String OUTPUT_FILENAME = "Simple.xlsx";
    private static final XMLEventFactory XEF = XMLEventFactory.newInstance();
    private static final XMLInputFactory XIF = XMLInputFactory.newInstance();
    private static final XMLOutputFactory XOF = XMLOutputFactory.newInstance();

    public static void main(String[] args) {
        try(ZipInputStream zis = new ZipInputStream(new FileInputStream(INPUT_FILENAME));
            ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(OUTPUT_FILENAME));) {
            Properties mappings = new Properties();
            mappings.load(OoXmlStrictConverter.class.getResourceAsStream("/ooxml-strict-mappings.properties"));
            ZipEntry ze;
            while((ze = zis.getNextEntry()) != null) {
                ZipEntry newZipEntry = new ZipEntry(ze.getName());
                zos.putNextEntry(newZipEntry);
                FilterInputStream filterIs = new FilterInputStream(zis) {
                    @Override
                    public void close() throws IOException {
                    }
                };
                FilterOutputStream filterOs = new FilterOutputStream(zos) {
                    @Override
                    public void close() throws IOException {
                    }
                };
                XMLEventReader xer = XIF.createXMLEventReader(filterIs);
                XMLEventWriter xew = XOF.createXMLEventWriter(filterOs);
                while(xer.hasNext()) {
                    XMLEvent xe = xer.nextEvent();
                    if(xe.isStartElement()) {
                        StartElement se = xe.asStartElement();
                        String namespaceUri = se.getName().getNamespaceURI();
                        if(namespaceUri != null && !namespaceUri.isEmpty()) {
                            String mappedUri = mappings.getProperty(namespaceUri);
                            if(mappedUri != null) {
                                xe = XEF.createStartElement(
                                        new QName(mappedUri, se.getName().getLocalPart()),
                                        se.getAttributes(), se.getNamespaces());
                            }
                        }
                    } else if(xe.isEndElement()) {
                        EndElement ee = xe.asEndElement();
                        String namespaceUri = ee.getName().getNamespaceURI();
                        if(namespaceUri != null && !namespaceUri.isEmpty()) {
                            String mappedUri = mappings.getProperty(namespaceUri);
                            if(mappedUri != null) {
                                xe = XEF.createEndElement(
                                        new QName(mappedUri, ee.getName().getLocalPart()),
                                        ee.getNamespaces());
                            }
                        }
                    }
                    xew.add(xe);
                }
                xer.close();
                xew.close();
                zis.closeEntry();
                zos.closeEntry();
            }
        } catch(Throwable t) {
            t.printStackTrace();
        }
    }
}
