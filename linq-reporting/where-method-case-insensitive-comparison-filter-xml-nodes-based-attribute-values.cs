using System;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Sample XML content with <item> elements.
        string sampleXml = @"
            <items>
                <item type='Apple' name='Red Delicious' />
                <item type='banana' name='Cavendish' />
                <item type='APPLE' name='Granny Smith' />
                <item type='orange' name='Navel' />
            </items>";

        // Add the XML as a custom XML part.
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add("myXmlPart", sampleXml);

        // Parse the XML content of the custom part.
        string xmlContent = Encoding.UTF8.GetString(xmlPart.Data);
        XDocument xDoc = XDocument.Parse(xmlContent);

        // Filter all <item> elements where the "type" attribute equals "apple",
        // using case‑insensitive comparison.
        var filteredItems = xDoc
            .Descendants("item")
            .Where(e => string.Equals(
                (string)e.Attribute("type"),
                "apple",
                StringComparison.OrdinalIgnoreCase));

        // Output the number of matching nodes.
        Console.WriteLine($"Found {filteredItems.Count()} <item> elements with type='apple' (case‑insensitive).");

        // Save the document (optional, demonstrates that the document is still valid).
        doc.Save("output.docx");
    }
}
