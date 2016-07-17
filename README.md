# ooxml-strict-converter
Early prototype code to take Strict OOXML files and save them using the more portable Transitional OOXML format. The Strict OOXML was introduced in Microsoft Office 2013. The aim is to provide a conversion utility that can be used to convert docs so that can be parsed using Apache POI.

When I have fixed a few issues, I will tidy up the code.

## Issues
* Handle Type attributes containing Strict OOXML formatted URIs, eg Relationship Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
