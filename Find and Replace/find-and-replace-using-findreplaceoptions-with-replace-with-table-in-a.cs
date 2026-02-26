using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Set up find‑replace options with a custom callback that will insert a table.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertTableHandler();

        // Replace the placeholder "[TABLE]" with an empty string; the callback will insert the table.
        doc.Range.Replace(new Regex(@"\[TABLE\]"), string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Callback that inserts a table at the location of the matched placeholder.
    private class InsertTableHandler : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Create a simple 2x2 table.
            Table table = new Table(args.MatchNode.Document);
            // Ensure the table has at least one row and cell.
            table.EnsureMinimum();

            // Populate the table with sample data.
            for (int rowIdx = 0; rowIdx < 2; rowIdx++)
            {
                Row row = new Row(args.MatchNode.Document);
                table.AppendChild(row);
                for (int colIdx = 0; colIdx < 2; colIdx++)
                {
                    Cell cell = new Cell(args.MatchNode.Document);
                    cell.AppendChild(new Paragraph(args.MatchNode.Document));
                    cell.FirstParagraph.AppendChild(new Run(args.MatchNode.Document,
                        $"R{rowIdx + 1}C{colIdx + 1}"));
                    row.AppendChild(cell);
                }
            }

            // Insert the table after the paragraph that contains the placeholder.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;
            CompositeNode parent = placeholderParagraph.ParentNode;
            parent.InsertAfter(table, placeholderParagraph);

            // Remove the placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip the default replacement since we have already handled it.
            return ReplaceAction.Skip;
        }
    }
}
