using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load a document that contains smart tags.
        Document doc = new Document("Smart tags.docx");

        // Get all smart tags in the document.
        SmartTag[] smartTags = doc.GetChildNodes(NodeType.SmartTag, true).OfType<SmartTag>().ToArray();

        // Ensure there is at least one smart tag.
        if (smartTags.Length == 0)
        {
            Console.WriteLine("No smart tags found.");
            return;
        }

        // Access the CustomXmlPropertyCollection of the first smart tag.
        CustomXmlPropertyCollection properties = smartTags[0].Properties;

        // Declare the iteration variable with an explicit type (CustomXmlProperty) in the foreach loop.
        foreach (CustomXmlProperty property in properties)
        {
            // Access strongly typed properties of CustomXmlProperty.
            Console.WriteLine($"Name: {property.Name}, Value: {property.Value}, Uri: {property.Uri}");
        }
    }
}
