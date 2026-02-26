using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOCM document.
        Document doc = new Document("Input.docm");

        // Iterate through all paragraphs in the document.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Identify paragraphs that belong to a list (numbered or bulleted).
            if (para.ListFormat.IsListItem && para.ListFormat.List != null)
            {
                // Enable the restart-at-each-section flag for the list.
                // When the document is saved, Aspose.Words will emit the <w:restartNum/> tag
                // for the list items, which is the required behavior for DOCM format.
                para.ListFormat.List.IsRestartAtEachSection = true;
            }
        }

        // Save the modified document back to DOCM format.
        doc.Save("Output.docm", SaveFormat.Docm);
    }
}
