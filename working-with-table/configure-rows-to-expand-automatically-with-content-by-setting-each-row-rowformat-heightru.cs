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

        // Start a table.
        Table table = builder.StartTable();

        // First row - set HeightRule to Auto.
        builder.RowFormat.HeightRule = HeightRule.Auto;
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second row - longer text to demonstrate auto expansion.
        builder.RowFormat.HeightRule = HeightRule.Auto;
        builder.InsertCell();
        builder.Write("Row 2, Cell 1 with longer text that should cause the row to expand automatically.");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Third row.
        builder.RowFormat.HeightRule = HeightRule.Auto;
        builder.InsertCell();
        builder.Write("Row 3, Cell 1");
        builder.InsertCell();
        builder.Write("Row 3, Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "RowsAutoHeight.docx";
        doc.Save(outputPath);

        // Verify the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Reload the document and validate each row's HeightRule is set to Auto.
        Document loadedDoc = new Document(outputPath);
        foreach (Row row in loadedDoc.GetChildNodes(NodeType.Row, true))
        {
            if (row.RowFormat.HeightRule != HeightRule.Auto)
                throw new Exception("A row's HeightRule is not set to Auto.");
        }
    }
}
