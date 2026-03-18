using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class RemoveQuoteParagraphs
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Iterate through all paragraphs in the main story.
        // Collect paragraphs to remove to avoid modifying the collection while iterating.
        var paragraphsToRemove = new System.Collections.Generic.List<Paragraph>();

        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Check if the paragraph's style identifier matches the target style (Quote).
            if (para.ParagraphFormat.Style != null &&
                para.ParagraphFormat.Style.StyleIdentifier == StyleIdentifier.Quote)
            {
                paragraphsToRemove.Add(para);
            }
        }

        // Remove the collected paragraphs from the document.
        foreach (Paragraph para in paragraphsToRemove)
        {
            para.Remove();
        }

        // Optionally clean up any now‑unused styles.
        doc.Cleanup();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
