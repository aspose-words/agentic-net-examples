using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;

class ReplacePlaceholderWithTable
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Set up find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertTableHandler();

        // Replace the placeholder "[TABLE]" with a table.
        doc.Range.Replace(new Regex(@"\[TABLE\]"), string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Callback that inserts a table at the location of each match.
    private class InsertTableHandler : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match is inside a paragraph; get that paragraph.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;

            // Create a simple 2x2 table.
            Table table = new Table(placeholderParagraph.Document);
            // Ensure the table has at least one row.
            table.EnsureMinimum();

            // First row.
            Row row1 = new Row(placeholderParagraph.Document);
            Cell cell11 = new Cell(placeholderParagraph.Document);
            cell11.AppendChild(new Paragraph(placeholderParagraph.Document));
            cell11.FirstParagraph.AppendChild(new Run(placeholderParagraph.Document, "Cell 1,1"));
            row1.AppendChild(cell11);
            Cell cell12 = new Cell(placeholderParagraph.Document);
            cell12.AppendChild(new Paragraph(placeholderParagraph.Document));
            cell12.FirstParagraph.AppendChild(new Run(placeholderParagraph.Document, "Cell 1,2"));
            row1.AppendChild(cell12);
            table.AppendChild(row1);

            // Second row.
            Row row2 = new Row(placeholderParagraph.Document);
            Cell cell21 = new Cell(placeholderParagraph.Document);
            cell21.AppendChild(new Paragraph(placeholderParagraph.Document));
            cell21.FirstParagraph.AppendChild(new Run(placeholderParagraph.Document, "Cell 2,1"));
            row2.AppendChild(cell21);
            Cell cell22 = new Cell(placeholderParagraph.Document);
            cell22.AppendChild(new Paragraph(placeholderParagraph.Document));
            cell22.FirstParagraph.AppendChild(new Run(placeholderParagraph.Document, "Cell 2,2"));
            row2.AppendChild(cell22);
            table.AppendChild(row2);

            // Insert the table after the placeholder paragraph.
            CompositeNode parent = placeholderParagraph.ParentNode;
            parent.InsertAfter(table, placeholderParagraph);

            // Remove the placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip the default replacement since we have already handled it.
            return ReplaceAction.Skip;
        }
    }
}
