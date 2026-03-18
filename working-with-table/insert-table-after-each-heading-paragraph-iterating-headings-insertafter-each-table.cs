using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableAfterHeadings
{
    static void Main()
    {
        // Load an existing document (replace with your file path).
        Document doc = new Document(@"Input.docx");

        // Collect all heading paragraphs (Heading1 through Heading9).
        List<Paragraph> headingParagraphs = new List<Paragraph>();
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;
            if (styleId >= StyleIdentifier.Heading1 && styleId <= StyleIdentifier.Heading9)
                headingParagraphs.Add(para);
        }

        // Insert a table after each heading. Process in reverse order to keep indexes stable.
        for (int i = headingParagraphs.Count - 1; i >= 0; i--)
        {
            Paragraph heading = headingParagraphs[i];

            // Create a simple 1x1 table.
            Table table = new Table(doc);
            table.EnsureMinimum(); // Guarantees at least one row and one cell.

            // Add some sample text to the first cell.
            Cell firstCell = table.FirstRow.FirstCell;
            firstCell.FirstParagraph.AppendChild(new Run(doc, "Table after heading"));

            // Insert the table immediately after the heading paragraph.
            // The heading's parent node is a CompositeNode (the Body).
            heading.ParentNode.InsertAfter(table, heading);
        }

        // Save the modified document.
        doc.Save(@"Output.docx");
    }
}
