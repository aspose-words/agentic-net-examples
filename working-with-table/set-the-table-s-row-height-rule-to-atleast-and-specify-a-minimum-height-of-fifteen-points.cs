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

        // Configure the next row: minimum height of 15 points and HeightRule.AtLeast.
        builder.RowFormat.Height = 15;
        builder.RowFormat.HeightRule = HeightRule.AtLeast;

        // Second row that will use the above height settings.
        builder.InsertCell();
        builder.Write("Second row, cell 1.");
        builder.InsertCell();
        builder.Write("Second row, cell 2.");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SetRowHeight.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
