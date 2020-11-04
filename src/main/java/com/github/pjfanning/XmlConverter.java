package com.github.pjfanning;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.Properties;
import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventFactory;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLEventWriter;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.Namespace;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

public class XmlConverter implements AutoCloseable {

    private static final XMLEventFactory XEF = XMLEventFactory.newInstance();
    private static final XMLInputFactory XIF = XMLInputFactory.newInstance();
    private static final XMLOutputFactory XOF = XMLOutputFactory.newInstance();
    private static final QName CONFORMANCE = new QName("conformance");
    private static final Properties mappings;

    static {
        XOF.setProperty(XMLOutputFactory.IS_REPAIRING_NAMESPACES, true);
        mappings = OoXmlStrictConverterUtils.readMappings();
    }

    private final XMLEventWriter xew;
    private final XMLEventReader xer;
    private int depth = 0;

    public XmlConverter(InputStream is, OutputStream os) throws XMLStreamException {
        this.xer = XIF.createXMLEventReader(is);
        this.xew = XOF.createXMLEventWriter(os);
    }

    public boolean convertNextElement() throws XMLStreamException {
        if (!xer.hasNext()) {
            return false;
        }

        XMLEvent xe = xer.nextEvent();
        if(xe.isStartElement()) {
            xew.add(convertStartElement(xe.asStartElement(), depth==0));
            depth++;
        } else if(xe.isEndElement()) {
            xew.add(convertEndElement(xe.asEndElement()));
            depth--;
        }

        xew.flush();

        return true;
    }

    @Override
    public void close() throws XMLStreamException {
        xer.close();
        xew.close();
    }

    private static StartElement convertStartElement(StartElement startElement, boolean root) {
        return  XEF.createStartElement(updateQName(startElement.getName()),
                processAttributes(startElement.getAttributes(), startElement.getName().getNamespaceURI(), root),
                processNamespaces(startElement.getNamespaces()));
    }

    private static EndElement convertEndElement(EndElement endElement) {
        return XEF.createEndElement(updateQName(endElement.getName()),
                processNamespaces(endElement.getNamespaces()));

    }

    private static QName updateQName(QName qn) {
        String namespaceUri = qn.getNamespaceURI();
        if(OoXmlStrictConverterUtils.isNotBlank(namespaceUri)) {
            String mappedUri = mappings.getProperty(namespaceUri);
            if(mappedUri != null) {
                qn = OoXmlStrictConverterUtils.isBlank(qn.getPrefix()) ? new QName(mappedUri, qn.getLocalPart())
                        : new QName(mappedUri, qn.getLocalPart(), qn.getPrefix());
            }
        }
        return qn;
    }

    private static Iterator<Attribute> processAttributes(final Iterator<Attribute> iter,
            final String elementNamespaceUri, final boolean rootElement) {
        ArrayList<Attribute> list = new ArrayList<>();
        while(iter.hasNext()) {
            Attribute att = iter.next();
            QName qn = updateQName(att.getName());
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

    private static Iterator<Namespace> processNamespaces(final Iterator<Namespace> iter) {
        ArrayList<Namespace> list = new ArrayList<>();
        while(iter.hasNext()) {
            Namespace ns = iter.next();
            if(!ns.isDefaultNamespaceDeclaration() && !mappings.containsKey(ns.getNamespaceURI())) {
                list.add(ns);
            }
        }
        return Collections.unmodifiableList(list).iterator();
    }

}
