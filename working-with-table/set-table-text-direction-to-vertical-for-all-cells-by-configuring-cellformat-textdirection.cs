using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the text orientation for all cells that will be created.
        // This setting is applied globally to the builder's CellFormat.
        builder.CellFormat.Orientation = TextOrientation.Upward;

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Verify that every cell in the table has the expected orientation.
        foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
        {
            if (cell.CellFormat.Orientation != TextOrientation.Upward)
                throw new InvalidOperationException("A cell does not have the expected vertical text orientation.");
        }

        // Save the document.
        const string outputPath = "TableTextDirection.docx";
        doc.Save(outputPath);

        // Simple confirmation (optional, does not require user input).
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
