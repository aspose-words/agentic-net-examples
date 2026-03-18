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
        // Load an existing Word document that contains a custom XML part.
        Document doc = new Document("input.docx");

        // Assume the document has at least one custom XML part.
        // Retrieve the first custom XML part.
        CustomXmlPart xmlPart = doc.CustomXmlParts[0];

        // Parse the XML content of the custom part.
        string xmlContent = Encoding.UTF8.GetString(xmlPart.Data);
        XDocument xDoc = XDocument.Parse(xmlContent);

        // Example: filter all <item> elements where the "type" attribute equals "apple",
        // using case‑insensitive comparison.
        var filteredItems = xDoc
            .Descendants("item")
            .Where(e => string.Equals(
                (string)e.Attribute("type"),
                "apple",
                StringComparison.OrdinalIgnoreCase));

        // Output the number of matching nodes (for demonstration purposes).
        Console.WriteLine($"Found {filteredItems.Count()} <item> elements with type='apple' (case‑insensitive).");

        // Optionally, you could map the filtered XML nodes back to structured document tags,
        // or modify the document based on the results here.

        // Save the (potentially modified) document.
        doc.Save("output.docx");
    }
}
