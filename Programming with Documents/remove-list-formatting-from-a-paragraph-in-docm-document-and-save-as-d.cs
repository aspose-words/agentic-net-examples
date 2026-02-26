using System;
using Aspose.Words;
using Aspose.Words.Lists;

class RemoveListFormatting
{
    static void Main()
    {
        // Load the macro-enabled DOCM document.
        Document doc = new Document("InputDocument.docm");

        // Iterate through all paragraphs in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph paragraph in paragraphs)
        {
            // If the paragraph is part of a list, remove its list formatting.
            if (paragraph.ListFormat.IsListItem)
                paragraph.ListFormat.RemoveNumbers();
        }

        // Optional: clean up any unused list definitions after removal.
        doc.Cleanup();

        // Save the modified document as a macro-enabled template (DOTM).
        doc.Save("OutputDocument.dotm", SaveFormat.Dotm);
    }
}
