using System;
using System.Linq;
using System.Xml.Linq;

class Program
{
    static void Main()
    {
        // Sample XML content.
        string xmlContent = @"
            <Root>
                <Item Id='1' Name='Alpha' />
                <Item Id='2' Name='Beta' />
                <Item Id='3' Name='Gamma' />
            </Root>";

        // Load the XML document from the string.
        XDocument xdoc = XDocument.Parse(xmlContent);

        // This will hold the concatenated attribute values.
        string composite = string.Empty;

        // Iterate over each <Item> element.
        foreach (XElement element in xdoc.Descendants("Item"))
        {
            // Concatenate all attribute values of the current element.
            string attrs = string.Concat(element.Attributes().Select(a => a.Value));

            // Append to the composite string, separating entries with a pipe for readability.
            composite = string.IsNullOrEmpty(composite) ? attrs : composite + "|" + attrs;
        }

        // Output the resulting composite string.
        Console.WriteLine("Composite attribute string:");
        Console.WriteLine(composite);
    }
}
