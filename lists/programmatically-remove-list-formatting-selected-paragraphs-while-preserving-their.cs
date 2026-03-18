using System;
using Aspose.Words;
using Aspose.Words.Lists;

class RemoveListFormatting
{
    static void Main()
    {
        // Load the existing document.
        Document doc = new Document("Input.docx");

        // Get all paragraphs in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        foreach (Paragraph para in paragraphs)
        {
            // Process only paragraphs that are part of a list.
            if (para.ListFormat.IsListItem)
            {
                // Preserve the current left indentation.
                double originalLeftIndent = para.ParagraphFormat.LeftIndent;

                // Remove the list numbering/bullet and reset the list level.
                para.ListFormat.RemoveNumbers();

                // Re‑apply the preserved indentation.
                para.ParagraphFormat.LeftIndent = originalLeftIndent;
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
