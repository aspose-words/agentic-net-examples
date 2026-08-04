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

        // First row.
        builder.InsertCell();
        builder.Write("First row, first cell.");
        builder.InsertCell();
        builder.Write("First row, second cell.");
        builder.EndRow();

        // Set height rule to AtLeast with a minimum height of 15 points for the next row.
        builder.RowFormat.Height = 15;
        builder.RowFormat.HeightRule = HeightRule.AtLeast;

        // Second row.
        builder.InsertCell();
        builder.Write("Second row, first cell.");
        builder.InsertCell();
        builder.Write("Second row, second cell.");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "TableRowHeight.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");
    }
}
