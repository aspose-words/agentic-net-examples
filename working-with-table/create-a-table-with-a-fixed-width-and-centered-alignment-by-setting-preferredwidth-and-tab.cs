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

        // Start a new table.
        Table table = builder.StartTable();

        // First row with two cells.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // The table now has at least one row, so we can safely set formatting.
        table.PreferredWidth = PreferredWidth.FromPoints(300);
        table.Alignment = TableAlignment.Center;

        // Second row with two cells.
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "TableFixedWidthCentered.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
