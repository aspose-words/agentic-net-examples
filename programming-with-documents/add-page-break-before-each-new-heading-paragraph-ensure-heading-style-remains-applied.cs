using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddPageBreakBeforeHeadings
{
    static void Main()
    {
        // Load an existing document (replace with your source file path).
        Document doc = new Document(@"C:\Docs\Input.docx");

        // Iterate through all paragraphs in the document.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Check if the paragraph uses a built‑in heading style.
            if (para.ParagraphFormat.IsHeading)
            {
                // Force a page break before this heading.
                para.ParagraphFormat.PageBreakBefore = true;

                // The heading style is not altered; it remains applied.
            }
        }

        // Save the modified document (replace with your desired output path).
        doc.Save(@"C:\Docs\Output.docx");
    }
}
