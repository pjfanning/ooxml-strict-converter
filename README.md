# ooxml-strict-converter
Early prototype code to take Strict OOXML files and save them using the more portable Transitional OOXML format. The ability to save in Strict OOXML format was introduced in Microsoft Office 2013. The aim is to provide a conversion utility that can be used to convert docs so that they can be parsed using Apache POI. https://bz.apache.org/bugzilla/show_bug.cgi?id=57699

The OOXML schemas are at http://www.ecma-international.org/publications/standards/Ecma-376.htm

When I have fixed a few issues, I will tidy up the code.

## Current Algorithm
* Produce new output file based on input file, both files are treated as zip files
* Takes each XML file inside the input zip and uses StAX parser to read events, transform them and output to new XML file
* Takes all known Strict OOXML namespaces and maps them to equivalent Transitional OOXML namespace
* Checks attribute values like Relationship Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" and modifies the value to be equivalent Transitional OOXML value
* In xl/workbook.xml, drop the conformance=“strict” attribute on the root element
