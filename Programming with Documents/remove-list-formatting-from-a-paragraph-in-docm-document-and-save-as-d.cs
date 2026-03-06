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
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // If the paragraph is part of a list, remove its list formatting.
            if (para.ListFormat.IsListItem)
                para.ListFormat.RemoveNumbers();
        }

        // Save the modified document as a macro-enabled template (DOTM).
        doc.Save("OutputDocument.dotm", SaveFormat.Dotm);
    }
}
