package com.example;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilterInputStream;
import java.io.FilterOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
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
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.Namespace;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

public class OoXmlStrictConverter {

    private static final String INPUT_FILENAME = "SimpleStrict.xlsx";
    private static final String OUTPUT_FILENAME = "Simple.xlsx";
    private static final XMLEventFactory XEF = XMLEventFactory.newInstance();
    private static final XMLInputFactory XIF = XMLInputFactory.newInstance();
    private static final XMLOutputFactory XOF = XMLOutputFactory.newInstance();

    public static void main(String[] args) {
        XOF.setProperty(XMLOutputFactory.IS_REPAIRING_NAMESPACES, true);
        try(ZipInputStream zis = new ZipInputStream(new FileInputStream(INPUT_FILENAME));
            ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(OUTPUT_FILENAME));) {
            Properties mappings = readMappings();
            System.out.println("loaded mappings entries=" + mappings.size());
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
                int depth = 0;
                while(xer.hasNext()) {
                    XMLEvent xe = xer.nextEvent();
                    if(xe.isStartElement()) {
                        StartElement se = xe.asStartElement();
                        xe = XEF.createStartElement(updateQName(se.getName(), mappings),
                                processAttributes(se.getAttributes(), mappings, se.getName().getNamespaceURI(), (depth == 0)),
                                processNamespaces(se.getNamespaces(), mappings));
                        depth++;
                    } else if(xe.isEndElement()) {
                        EndElement ee = xe.asEndElement();
                        xe = XEF.createEndElement(updateQName(ee.getName(), mappings),
                                processNamespaces(ee.getNamespaces(), mappings));
                        depth--;
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

    private static final QName CONFORMANCE = new QName("conformance");
    
    private static Iterator<Attribute> processAttributes(final Iterator<Attribute> iter,
            final Properties mappings, final String elementNamespaceUri, final boolean rootElement) {
        ArrayList<Attribute> list = new ArrayList<>();
        while(iter.hasNext()) {
            Attribute att = iter.next();
            QName qn = updateQName(att.getName(), mappings);
            if(rootElement && mappings.containsKey(elementNamespaceUri) && att.getName().equals(CONFORMANCE)) {
                //drop attribute
            } else {
                String newValue = att.getValue();
                for(String key : mappings.stringPropertyNames()) {
                    if(att.getValue().startsWith(key)) {
                        newValue = att.getValue().replace(key, mappings.getProperty(key));
                        break;
                    }
                }
                list.add(XEF.createAttribute(qn, newValue));
            }
        }
        return Collections.unmodifiableList(list).iterator();
    }

    private static Iterator<Namespace> processNamespaces(final Iterator<Namespace> iter,
            final Properties mappings) {
        ArrayList<Namespace> list = new ArrayList<>();
        while(iter.hasNext()) {
            Namespace ns = iter.next();
            if(!ns.isDefaultNamespaceDeclaration() && !mappings.containsKey(ns.getNamespaceURI())) {
                list.add(ns);
            }
        }
        return Collections.unmodifiableList(list).iterator();
    }

    private static QName updateQName(QName qn, Properties mappings) {
        String namespaceUri = qn.getNamespaceURI();
        if(isNotBlank(namespaceUri)) {
            String mappedUri = mappings.getProperty(namespaceUri);
            if(mappedUri != null) {
                qn = isBlank(qn.getPrefix()) ? new QName(mappedUri, qn.getLocalPart())
                        : new QName(mappedUri, qn.getLocalPart(), qn.getPrefix());
            }
        }
        return qn;
    }
    
    private static Properties readMappings() throws IOException {
        Properties props = new Properties();
        try(InputStream is = OoXmlStrictConverter.class.getResourceAsStream("/ooxml-strict-mappings.properties");
            BufferedReader reader = new BufferedReader(new InputStreamReader(is, "ISO-8859-1"))) {
            String line;
            while((line = reader.readLine()) != null) {
                String[] vals = line.split("=");
                if(vals.length >= 2) {
                    props.setProperty(vals[0], vals[1]);
                } else if(vals.length == 1) {
                    props.setProperty(vals[0], "");
                }

            }
        }
        return props;
    }
    
    private static boolean isBlank(final String str) {
        return str == null || str.trim().length() == 0;
    }

    private static boolean isNotBlank(final String str) {
        return !isBlank(str);
    }
}
