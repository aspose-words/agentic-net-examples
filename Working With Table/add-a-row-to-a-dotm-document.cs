using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing DOTM template.
        Document doc = new Document("Template.dotm");

        // Locate the first table in the document (adjust the index if needed).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("No table found in the document.");

        // Create a new row that belongs to the same document.
        Row newRow = new Row(doc);

        // Ensure the row has at least one cell (required for a valid row).
        newRow.EnsureMinimum();

        // Add some sample text to the first cell of the new row.
        newRow.FirstCell.FirstParagraph.AppendChild(new Run(doc, "New row content"));

        // Append the new row to the end of the table.
        table.AppendChild(newRow);

        // Save the modified document back to DOTM format.
        doc.Save("ModifiedTemplate.dotm");
    }
}
