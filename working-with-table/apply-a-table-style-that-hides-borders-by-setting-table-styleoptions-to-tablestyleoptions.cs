using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2‑row, 1‑column table.
        builder.StartTable();

        builder.InsertCell();
        builder.Write("First row");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Second row");
        builder.EndRow();

        builder.EndTable();

        // Get the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a built‑in style (optional, here we use TableGrid).
        table.StyleIdentifier = StyleIdentifier.TableGrid;

        // Hide all borders by clearing the border formatting.
        table.ClearBorders();

        // Save the document.
        string outputFile = "TableNoBorders.docx";
        doc.Save(outputFile);

        // Verify that the file was saved.
        if (!File.Exists(outputFile))
            throw new Exception("The output document was not created.");
    }
}
