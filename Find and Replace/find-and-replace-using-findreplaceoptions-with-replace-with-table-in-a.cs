using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace to use a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new TableReplacingHandler();

        // Replace the placeholder "[TABLE]" with a table.
        doc.Range.Replace(new Regex(@"\[TABLE\]"), string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Callback that inserts a table where the placeholder was found.
    private class TableReplacingHandler : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The document that owns the match. Cast from DocumentBase to Document.
            Document ownerDoc = (Document)args.MatchNode.Document;

            // Build a simple 2x2 table.
            Table table = new Table(ownerDoc);
            for (int rowIdx = 0; rowIdx < 2; rowIdx++)
            {
                Row row = new Row(ownerDoc);
                table.AppendChild(row);
                for (int colIdx = 0; colIdx < 2; colIdx++)
                {
                    Cell cell = new Cell(ownerDoc);
                    cell.AppendChild(new Paragraph(ownerDoc));
                    cell.FirstParagraph.AppendChild(new Run(ownerDoc, $"R{rowIdx + 1}C{colIdx + 1}"));
                    row.AppendChild(cell);
                }
            }

            // Insert the table after the paragraph that contains the placeholder.
            // The match node is a Run; its parent is the Paragraph that holds the placeholder.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;
            placeholderParagraph.ParentNode.InsertAfter(table, placeholderParagraph);

            // Remove the placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip the default text replacement (we already handled it).
            return ReplaceAction.Skip;
        }
    }
}
