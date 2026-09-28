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

        // Build a simple table with one row and one cell.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the first table and its first row.
        Table table = doc.FirstSection.Body.Tables[0];
        Row row = table.Rows[0];

        // Set the row height rule to Auto (do not specify an explicit height).
        row.RowFormat.HeightRule = HeightRule.Auto;

        // Save the document to a file.
        string outputPath = "TableRowHeightAuto.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Optionally, inform that the process completed successfully.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
