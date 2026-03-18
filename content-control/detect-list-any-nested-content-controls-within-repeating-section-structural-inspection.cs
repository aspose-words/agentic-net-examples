using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the Word document that contains repeating sections.
        // Replace the path with the actual location of your document.
        Document doc = new Document("InputDocument.docx");

        // Retrieve all structured document tags (content controls) in the document.
        NodeCollection allSdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        // Iterate through each content control that is a repeating section.
        foreach (StructuredDocumentTag repeatingSection in allSdtNodes
                     .OfType<StructuredDocumentTag>()
                     .Where(sdt => sdt.SdtType == SdtType.RepeatingSection))
        {
            Console.WriteLine($"Repeating Section found. Title: '{repeatingSection.Title}', Tag: '{repeatingSection.Tag}'");

            // Get all nested content controls inside this repeating section (deep search).
            // The repeating section itself is also a StructuredDocumentTag, so we filter it out.
            NodeCollection nestedControls = repeatingSection.GetChildNodes(NodeType.StructuredDocumentTag, true);

            foreach (StructuredDocumentTag nestedControl in nestedControls
                         .OfType<StructuredDocumentTag>()
                         .Where(sdt => sdt != repeatingSection))
            {
                // Output useful information about each nested content control.
                Console.WriteLine($"  Nested Control - Type: {nestedControl.SdtType}, Title: '{nestedControl.Title}', Tag: '{nestedControl.Tag}'");
            }

            Console.WriteLine(); // Blank line for readability between sections.
        }

        // Optionally, save the document after inspection (no modifications made here).
        doc.Save("OutputDocument.docx");
    }
}
