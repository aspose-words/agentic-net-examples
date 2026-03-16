using System;
using Aspose.Words;
using Aspose.Words.Notes;

class Program
{
    static void Main()
    {
        // Load the original DOCX.
        Document doc = new Document("Input.docx");

        // Clone the document to preserve the original layout for comparison.
        Document before = (Document)doc.Clone(true);
        before.UpdatePageLayout();
        int pagesBefore = before.PageCount;

        // Disable automatic hyphenation for footnote paragraphs only.
        foreach (Footnote footnote in doc.GetChildNodes(NodeType.Footnote, true))
        {
            foreach (Paragraph para in footnote.GetChildNodes(NodeType.Paragraph, true))
            {
                // Suppress hyphenation for this paragraph.
                para.ParagraphFormat.SuppressAutoHyphens = true;
            }
        }

        // Rebuild the layout after the changes.
        doc.UpdatePageLayout();
        int pagesAfter = doc.PageCount;

        // Save the modified document.
        doc.Save("Output.docx");

        // Output the comparison of footnote layout (page count) before and after.
        Console.WriteLine($"Pages before disabling footnote hyphenation: {pagesBefore}");
        Console.WriteLine($"Pages after disabling footnote hyphenation: {pagesAfter}");
    }
}
