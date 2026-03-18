using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source XML file.
        string xmlPath = "input.xml";

        // Load the XML document using LINQ to XML.
        XDocument xdoc = XDocument.Load(xmlPath);

        // This will hold the concatenated attribute values.
        string composite = string.Empty;

        // Iterate over each <Item> element (replace "Item" with your element name).
        foreach (XElement element in xdoc.Descendants("Item"))
        {
            // Concatenate all attribute values of the current element.
            string attrs = string.Concat(element.Attributes().Select(a => a.Value));

            // Append to the composite string, separating entries with a pipe for readability.
            composite = string.IsNullOrEmpty(composite) ? attrs : composite + "|" + attrs;
        }

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the resulting composite string into the document.
        builder.Writeln("Composite attribute string:");
        builder.Writeln(composite);

        // Save the document to disk.
        doc.Save("Output.docx");
    }
}
