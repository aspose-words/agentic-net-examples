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

        // First row with default formatting.
        builder.InsertCell();
        builder.Write("First row, cell 1.");
        builder.InsertCell();
        builder.Write("First row, cell 2.");
        builder.EndRow();

        // Configure the row height rule to AtLeast and set a minimum height of 15 points.
        builder.RowFormat.Height = 15;
        builder.RowFormat.HeightRule = HeightRule.AtLeast;

        // Second row will inherit the above settings.
        builder.InsertCell();
        builder.Write("Second row, cell 1.");
        builder.InsertCell();
        builder.Write("Second row, cell 2.");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableRowHeight.docx");
        doc.Save(outputPath);
    }
}
